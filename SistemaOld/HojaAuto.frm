VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgHojaAuto 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Hojas de Produccion"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   3135
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Imprime 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2160
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Hojas a imprimir"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
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
End
Attribute VB_Name = "PrgHojaAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstComposicion As Recordset
Dim spComposision As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPrecio As Recordset
Dim spPrecio As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstSolHoja As Recordset
Dim spSolHoja As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim XParam As String
Dim Vector(1000) As String
Dim Datos(100, 10) As String
Private xLote(100, 7) As String
Private Auxiliar(100, 7) As String
Dim Impre(3, 2) As Double
Dim WSolHoja As String
Dim WHoja As String
Dim WFecha As String
Dim WProducto As String
Dim WTerorico As String
Dim Lugar As Integer
Dim Cantidad As String
Dim WTeorico As String

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Lugar = 0

    spSolHoja = "ListaSolHojaTotalListado"
    Set rstSolHoja = db.OpenRecordset(spSolHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolHoja.RecordCount > 0 Then
    
    With rstSolHoja
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                If rstSolHoja!Renglon = 1 Then
                    Entra = "S"
                    For XDa = 1 To Lugar
                        If Vector(Lugar) = rstSolHoja!Hoja Then
                            Entra = "N"
                            Exit For
                        End If
                    Next XDa
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar) = rstSolHoja!Hoja
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

    If Lugar > 0 Then
        PrgHojaAuto.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgHojaAuto.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()

    For WWCicla = 1 To Lugar
    
        m$ = "Coloque la Hoja de Produccion en la Impresora"
        ca% = MsgBox(m$, 0, "Impresion de Hoja de Produccion")
    
        WSolHoja = Vector(WWCicla)
        Call Proceso
        Call Impresion
        Call Grabacion
        WSolHoja = Vector(WWCicla)
        WMarca = "X"
            
        XParam = "'" + WSolHoja + "','" _
                     + WMarca + "'"
                                           
        spSolHoja = "ModificaSolHojaImpresion " + XParam
        Set rstSolHoja = db.OpenRecordset(spSolHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    Next WWCicla
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgHojaAuto.Hide
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Sub Impresion()

    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            WHoja = rstHoja!Hoja + 1
        End With
        rstHoja.Close
            Else
        WHoja = "1"
    End If

    If WEmpresa <> 9 Then
        Open "lpt1" For Output As #1
        Rem Open "hoja.txt" For Output As #1
            Else
        Open "hoja.txt" For Output As #1
    End If

    Print #1, Chr$(27) + Chr$(71)
    Print #1,
    Print #1, Chr$(18)

    Print #1, Tab(15); Left$(WProducto, 2);
    Select Case Val(WEmpresa)
        Case 1
            Print #1, Tab(70); "SI"
        Case 2
            Print #1, Tab(70); "PI"
        Case 3
            Print #1, Tab(70); "SII"
        Case 4
            Print #1, Tab(70); "PII"
        Case 5
            Print #1, Tab(70); "SIII"
        Case 6
            Print #1, Tab(70); "SIV"
        Case 7
            Print #1, Tab(70); "SV"
        Case 8
            Print #1, Tab(70); "PIII"
        Case Else
    End Select

    Print #1, Tab(1); WFecha;
    Print #1, Tab(12); Alinea("#####", Mid$(WProducto, 4, 5));
    Print #1, "/"; Right$(WProducto, 3);
    Print #1, Tab(26); Chr$(14); Alinea("######", WHoja)

    Print #1,
    Print #1,

    Linea = 0
        
    For A = 1 To 40
        
        Tipo = Datos(A, 1)
        Terminado = Datos(A, 2)
        Articulo = Datos(A, 3)
        Cantidad = Datos(A, 4)
        If Val(xLote(A, 1)) <> 0 Then
            Pasa = "N"
                Else
            Pasa = "S"
        End If
        
        If Pasa = "S" Then
                 
            If Tipo = "M" Then
                    
                Rem PROCESA LOS LAUDOS
    
                Erase Impre
                Xlugar = 0
                XCanti = Val(Cantidad)
    
                XParam = "'" + Articulo + "','" _
                             + Articulo + "'"
                spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
    
                    With rstLaudo
    
                        .MoveFirst
                
                        If .NoMatch = False Then
                            Do
            
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                                If rstLaudo!Saldo <> 0 Then
                                    If rstLaudo!Articulo = Articulo Then
                                        If Xlugar < 3 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstLaudo!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstLaudo!Laudo
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                    Else
                                                Impre(Xlugar, 1) = rstLaudo!Laudo
                                                Impre(Xlugar, 2) = rstLaudo!Saldo
                                                XCanti = XCanti - rstLaudo!Saldo
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
                    rstLaudo.Close
                End If
                        
                XParam = "'" + Articulo + "','" _
                             + Articulo + "'"
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
                
                                If rstMovguia!Saldo <> 0 Then
                                    If rstMovguia!Articulo = Articulo Then
                                        If Xlugar < 3 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstMovguia!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                    Else
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = rstMovguia!Saldo
                                                XCanti = XCanti - rstMovguia!Saldo
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
                        
                Linea = Linea + 1

                Print #1, Tab(6); Left$(Articulo, 2);
                Print #1, Tab(11); Mid$(Articulo, 4, 3);
                Print #1, "-";
                Print #1, Right$(Articulo, 3);
                If Val(WTeorico) < 100 Then
                    Print #1, Tab(20); Alinea("###.##", Cantidad);
                        Else
                    Print #1, Tab(20); Alinea("####.#", Cantidad);
                End If
                        
                If Impre(1, 2) <> 0 Then
                    If Impre(1, 2) < 100 Then
                        Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                            Else
                        Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                    End If
                End If
                If Impre(1, 1) <> 0 Then
                    Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                End If
                
                If Impre(2, 2) <> 0 Then
                    If Impre(2, 2) < 100 Then
                        Print #1, Tab(41); Alinea("###.##", Str$(Impre(2, 2)));
                            Else
                        Print #1, Tab(41); Alinea("####.#", Str$(Impre(2, 2)));
                    End If
                End If
                If Impre(2, 1) <> 0 Then
                    Print #1, Tab(48); Alinea("######", Str$(Impre(2, 1)));
                End If
                        
                If Impre(3, 2) <> 0 Then
                    If Impre(3, 2) < 100 Then
                        Print #1, Tab(55); Alinea("###.##", Str$(Impre(3, 2)));
                            Else
                        Print #1, Tab(55); Alinea("####.#", Str$(Impre(3, 2)));
                    End If
                End If
                If Impre(3, 1) <> 0 Then
                    Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                End If
                            
                Print #1,
                Print #1,

            End If

            If Tipo = "T" Then
                    
                Erase Impre
                Xlugar = 0
                XCanti = Val(Cantidad)
    
                XParam = "'" + Terminado + "','" _
                             + Terminado + "'"
                spHoja = "ListaHojaProductoDesdeHasta" + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
    
                    With rstHoja
    
                        .MoveFirst
                
                        If .NoMatch = False Then
                            Do
            
                                If .EOF = True Then
                                    Exit Do
                                End If
                
                                If rstHoja!Saldo <> 0 And rstHoja!Renglon = 1 Then
                                    If rstHoja!Producto = Terminado Then
                                        If Xlugar < 3 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstHoja!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstHoja!Hoja
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                    Else
                                                Impre(Xlugar, 1) = rstHoja!Hoja
                                                Impre(Xlugar, 2) = rstHoja!Saldo
                                                XCanti = XCanti - rstHoja!Saldo
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
                    rstHoja.Close
                End If
                        
                XParam = "'" + Terminado + "','" _
                             + Terminado + "'"
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
                
                                If rstMovguia!Saldo <> 0 Then
                                    If rstMovguia!Terminado = Terminado Then
                                        If Xlugar < 3 And XCanti > 0 Then
                                            Xlugar = Xlugar + 1
                                            If rstMovguia!Saldo > XCanti Then
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = XCanti
                                                XCanti = 0
                                                            Else
                                                Impre(Xlugar, 1) = rstMovguia!Lote
                                                Impre(Xlugar, 2) = rstMovguia!Saldo
                                                XCanti = XCanti - rstMovguia!Saldo
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
                        

                Linea = Linea + 1

                Print #1, Tab(6); Left$(Terminado, 2);
                Print #1, Tab(11); Mid$(Terminado, 4, 5);
                Print #1, "-";
                Print #1, Right$(Terminado, 3);
                If Val(WTeorico) < 100 Then
                    Print #1, Tab(20); Alinea("###.##", Cantidad);
                        Else
                    Print #1, Tab(20); Alinea("####.#", Cantidad);
                End If
                        
                If Impre(1, 2) <> 0 Then
                    If Impre(1, 2) < 100 Then
                        Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                            Else
                        Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                    End If
                End If
                If Impre(1, 1) <> 0 Then
                    Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                End If
                        
                If Impre(2, 2) <> 0 Then
                    If Impre(2, 2) < 100 Then
                        Print #1, Tab(40); Alinea("###.##", Str$(Impre(2, 2)));
                            Else
                        Print #1, Tab(40); Alinea("####.#", Str$(Impre(2, 2)));
                    End If
                End If
                If Impre(2, 1) <> 0 Then
                    Print #1, Tab(46); Alinea("######", Str$(Impre(2, 1)));
                End If
                        
                If Impre(3, 2) <> 0 Then
                    If Impre(3, 2) < 100 Then
                        Print #1, Tab(54); Alinea("###.##", Str$(Impre(3, 2)));
                            Else
                        Print #1, Tab(54); Alinea("####.#", Str$(Impre(3, 2)));
                    End If
                End If
                If Impre(3, 1) <> 0 Then
                    Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                End If
                        
                Print #1,
                Print #1,

            End If

        End If
        
        If Pasa = "N" Then
                 
            If Tipo = "M" Then
                    
                Rem PROCESA LOS LAUDOS
    
                Erase Impre
                Impre(1, 1) = xLote(A, 1)
                Impre(1, 2) = xLote(A, 2)
                Impre(2, 1) = xLote(A, 3)
                Impre(2, 2) = xLote(A, 4)
                Impre(3, 1) = xLote(A, 5)
                Impre(3, 2) = xLote(A, 6)
                
                Linea = Linea + 1

                Print #1, Tab(6); Left$(Articulo, 2);
                Print #1, Tab(11); Mid$(Articulo, 4, 3);
                Print #1, "-";
                Print #1, Right$(Articulo, 3);
                If Val(WTeorico) < 100 Then
                    Print #1, Tab(20); Alinea("###.##", Cantidad);
                        Else
                    Print #1, Tab(20); Alinea("####.#", Cantidad);
                End If
                        
                If Impre(1, 2) <> 0 Then
                    If Impre(1, 2) < 100 Then
                        Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                            Else
                        Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                    End If
                End If
                If Impre(1, 1) <> 0 Then
                    Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                End If
                
                If Impre(2, 2) <> 0 Then
                    If Impre(2, 2) < 100 Then
                        Print #1, Tab(41); Alinea("###.##", Str$(Impre(2, 2)));
                            Else
                        Print #1, Tab(41); Alinea("####.#", Str$(Impre(2, 2)));
                    End If
                End If
                If Impre(2, 1) <> 0 Then
                    Print #1, Tab(48); Alinea("######", Str$(Impre(2, 1)));
                End If
                        
                If Impre(3, 2) <> 0 Then
                    If Impre(3, 2) < 100 Then
                        Print #1, Tab(55); Alinea("###.##", Str$(Impre(3, 2)));
                            Else
                        Print #1, Tab(55); Alinea("####.#", Str$(Impre(3, 2)));
                    End If
                End If
                If Impre(3, 1) <> 0 Then
                    Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                End If
                            
                Print #1,
                Print #1,

            End If

            If Tipo = "T" Then
                    
                Erase Impre
                
                Erase Impre
                Impre(1, 1) = xLote(A, 1)
                Impre(1, 2) = xLote(A, 2)
                Impre(2, 1) = xLote(A, 3)
                Impre(2, 2) = xLote(A, 4)
                Impre(3, 1) = xLote(A, 5)
                Impre(3, 2) = xLote(A, 6)

                Linea = Linea + 1

                Print #1, Tab(6); Left$(Terminado, 2);
                Print #1, Tab(11); Mid$(Terminado, 4, 5);
                Print #1, "-";
                Print #1, Right$(Terminado, 3);
                If Val(WTeorico) < 100 Then
                    Print #1, Tab(20); Alinea("###.##", Cantidad);
                        Else
                    Print #1, Tab(20); Alinea("####.#", Cantidad);
                End If
                        
                If Impre(1, 2) <> 0 Then
                    If Impre(1, 2) < 100 Then
                        Print #1, Tab(27); Alinea("###.##", Str$(Impre(1, 2)));
                            Else
                        Print #1, Tab(27); Alinea("####.#", Str$(Impre(1, 2)));
                    End If
                End If
                If Impre(1, 1) <> 0 Then
                    Print #1, Tab(34); Alinea("######", Str$(Impre(1, 1)));
                End If
                        
                If Impre(2, 2) <> 0 Then
                    If Impre(2, 2) < 100 Then
                        Print #1, Tab(40); Alinea("###.##", Str$(Impre(2, 2)));
                            Else
                        Print #1, Tab(40); Alinea("####.#", Str$(Impre(2, 2)));
                    End If
                End If
                If Impre(2, 1) <> 0 Then
                    Print #1, Tab(46); Alinea("######", Str$(Impre(2, 1)));
                End If
                        
                If Impre(3, 2) <> 0 Then
                    If Impre(3, 2) < 100 Then
                        Print #1, Tab(54); Alinea("###.##", Str$(Impre(3, 2)));
                            Else
                        Print #1, Tab(54); Alinea("####.#", Str$(Impre(3, 2)));
                    End If
                End If
                If Impre(3, 1) <> 0 Then
                    Print #1, Tab(62); Alinea("######", Str$(Impre(3, 1)));
                End If
                        
                Print #1,
                Print #1,

            End If

        End If
        
        
    Next A

    For Ciclo = Linea To 14
        Print #1,
        Print #1,
    Next Ciclo

    Print #1, Tab(20); Alinea("####.#", WTeorico)

    Print #1,
    Print #1, Chr$(27) + Chr$(72)
    Print #1, Chr$(12)
        
    Close #1

End Sub

Private Sub Proceso()

    Erase Datos
    Xlugar = 0

    spSolHoja = "ListaSolHoja " + "'" + WSolHoja + "'"
    Set rstSolHoja = db.OpenRecordset(spSolHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolHoja.RecordCount > 0 Then
        With rstSolHoja
            .MoveFirst
            Do
                If .EOF = False Then
            
                    Xlugar = Xlugar + 1
                    
                    WFecha = rstSolHoja!Fecha
                    WProducto = rstSolHoja!Producto
                    WTeorico = rstSolHoja!Teorico
                    
                    Datos(Xlugar, 1) = rstSolHoja!Tipo
                    Datos(Xlugar, 2) = rstSolHoja!Terminado
                    Datos(Xlugar, 3) = rstSolHoja!Articulo
                    Datos(Xlugar, 4) = Str$(rstSolHoja!Cantidad)
                    
                    xLote(Xlugar, 1) = IIf(IsNull(rstSolHoja!lote1), "", rstSolHoja!lote1)
                    xLote(Xlugar, 2) = IIf(IsNull(rstSolHoja!Canti1), "", rstSolHoja!Canti1)
                    xLote(Xlugar, 3) = IIf(IsNull(rstSolHoja!lote2), "", rstSolHoja!lote2)
                    xLote(Xlugar, 4) = IIf(IsNull(rstSolHoja!Canti2), "", rstSolHoja!Canti2)
                    xLote(Xlugar, 5) = IIf(IsNull(rstSolHoja!lote3), "", rstSolHoja!lote3)
                    xLote(Xlugar, 6) = IIf(IsNull(rstSolHoja!Canti3), "", rstSolHoja!Canti3)
                    xLote(Xlugar, 7) = ""
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstSolHoja.Close
    End If
    

End Sub

Private Sub Grabacion()

    Renglon = 0
    
    spHoja = "ListaHojaNumero"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveLast
            WHoja = rstHoja!Hoja + 1
        End With
        rstHoja.Close
            Else
        WHoja = "1"
    End If

    For A = 1 To 40
        
        Tipo = Datos(A, 1)
        Terminado = Datos(A, 2)
        Articulo = Datos(A, 3)
        Cantidad = Datos(A, 4)
                    
        If Articulo <> "" Then
                        
            Renglon = Renglon + 1
            Auxi = Str$(Renglon)
            Call Ceros(Auxi, 2)
                        
            Auxi1 = Str$(WHoja)
            Call Ceros(Auxi1, 6)
                    
            WClave = Auxi1 + Auxi
            WRenglon = Str$(Renglon)
            WReal = ""
            WFechaing = "  /  /    "
            WFechaingord = Right$(WFechaing, 4) + Mid$(WFechaing, 4, 2) + Left$(WFechaing, 2)
            WTipo = Tipo
            WArticulo = Articulo
            WTerminado = Terminado
            WCantidad = Cantidad
            WLote = ""
            WDate = Date$
            WImporte = ""
            WMarca = ""
            WSaldo = "0"
            WLote1 = "0"
            WLote2 = "0"
            Wlote3 = "0"
            WCanti1 = "0"
            WCanti2 = "0"
            WCanti3 = "0"
            WCosto1 = "0"
            WCosto2 = "0"
            WCosto3 = "0"
                    
            XParam = "'" + WClave + "','" _
                         + WHoja + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WProducto + "','" _
                         + WCantidad + "','" _
                         + WTipo + "','" _
                         + WLote + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WTeorico + "','" _
                         + WReal + "','" _
                         + WFechaing + "','" _
                         + WFechaingord + "','" _
                         + WDate + "','" _
                         + WImporte + "','" _
                         + WMarca + "','" _
                         + WSaldo + "','" _
                         + WLote1 + "','" + WCanti1 + "','" _
                         + WLote2 + "','" + WCanti2 + "','" _
                         + Wlote3 + "','" + Wlote3 + "','" _
                         + WCosto1 + "','" _
                         + WCosto2 + "','" _
                         + WCosto3 + "'"
                                           
            spHoja = "AltaHoja " + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
            Auxiliar(Renglon, 1) = WProducto
            Auxiliar(Renglon, 2) = WTerminado
            Auxiliar(Renglon, 3) = WArticulo
            Auxiliar(Renglon, 4) = WCantidad
            Auxiliar(Renglon, 5) = ""
            Auxiliar(Renglon, 6) = WTeorico
            Auxiliar(Renglon, 7) = WTipo
            
        End If
                        
    Next A
    
    WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
    XParam = "'" + WHoja + "','" _
                 + WFechaord + "'"
    Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    For Da = 1 To Renglon
    
        Producto = Auxiliar(Da, 1)
        Terminado = Auxiliar(Da, 2)
        Articulo = Auxiliar(Da, 3)
        Cantidad = Auxiliar(Da, 4)
        Real = Auxiliar(Da, 5)
        Teorico = Auxiliar(Da, 6)
        Tipo = Auxiliar(Da, 7)
        
        If Da = 1 Then
        
            spTerminado = "ConsultaTerminado " + "'" + Producto + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                WCodigo = rstTerminado!Codigo
                WProceso = Str$(rstTerminado!Proceso + Val(Teorico))
                WEntradas = Str$(rstTerminado!Entradas)
            End If
            WDate = Date$
            rstTerminado.Close
                    
            XParam = "'" + WCodigo + "','" _
                         + WEntradas + "','" _
                         + WProceso + "','" _
                         + WDate + "'"
                                           
            spTerminado = "ModificaTerminadoHoja " + XParam
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        End If
                
        Select Case Tipo
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WCodigo = rstArticulo!Codigo
                    WSalidas = Str$(rstArticulo!Salidas + Val(Cantidad))
                    WDate = Date$
                    rstArticulo.Close
                    XParam = "'" + WCodigo + "','" _
                                 + WSalidas + "','" _
                                 + WDate + "'"
                                                
                    spArticulo = "ModificaArticuloSalidas " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                End If
                                            
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + Terminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WCodigo = rstTerminado!Codigo
                    WSalidas = Str$(rstTerminado!Salidas + Val(Cantidad))
                    WDate = Date$
                    rstTerminado.Close
                            
                    XParam = "'" + WCodigo + "','" _
                                 + WSalidas + "','" _
                                 + WDate + "'"
                                            
                    spTerminado = "ModificaTerminadoSalidas " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                End If
                    
            Case Else
        End Select
        
    Next Da

End Sub
