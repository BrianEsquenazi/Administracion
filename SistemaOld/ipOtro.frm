VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgIpOtro 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos"
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
         Caption         =   "Pedidos a imprimir"
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
Attribute VB_Name = "PrgIpOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String
Dim Vector(1000) As String
Dim Datos(100, 13) As String
Dim WPedido As String
Dim WCliente As String
Dim WPago As String
Dim WDirentrega As String
Dim WObservaciones As String
Dim WDespago As String
Dim WFecha As String
Dim WFecEntrega As String
Dim WVersion As String
Dim WTipoped As String
Dim Lugar As Integer
Dim WEnvase(10) As String
Dim XEnvase(100, 6) As String
Dim Auxiliar(1000, 2) As String
Dim WImpre(10) As String

Private Sub Acepta_Click()

    WEmpresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    WEnvase(1) = 20
    WEnvase(2) = 21
    WEnvase(3) = 22
    WEnvase(4) = 23
    WEnvase(5) = 24
    WEnvase(6) = 25
    WEnvase(7) = 26
    WEnvase(8) = 30
    WEnvase(9) = 28

    For Cicla = 1 To 9
        spEnvase = "ConsultaEnvases " + "'" + WEnvase(Cicla) + "'"
        Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnvase.RecordCount > 0 Then
            WImpre(Cicla) = Left$(rstEnvase!Abreviatura, 7)
            rstEnvase.Close
                    Else
            WImpre(Cicla) = ""
        End If
    Next Cicla

    Erase XEnvase

    Lugar = 0

    spPedido = "ListaPedidoTotalListado1"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
    
    With rstPedido
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Rem If rstPedido!Renglon = 1 Then
                
                    If rstPedido!Autorizo = "X" Then
                    
                        If rstPedido!Impresion1 = "N" Then
                        
                            Entra = "S"
                            
                            For XDa = 1 To Lugar
                                If Vector(Lugar) = rstPedido!Pedido Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next XDa
                                
                            If Entra = "S" Then
                                Lugar = Lugar + 1
                                Vector(Lugar) = rstPedido!Pedido
                            End If
                            
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
    
    End If

    If Lugar > 0 Then
        PrgIpOtro.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgIpOtro.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If
    
End Sub

Private Sub Imprime_Click()
    If Val(WEmpresa) = 1 Then
        Rem Open "dada.txt" For Output As #1
        Open "lpt2" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "12" + Chr$(72);
            Else
        Rem Open "dada.txt" For Output As #1
        Open "lpt2" For Output As #1
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "3" + Chr$(65);
        Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "70" + Chr$(70)
        Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "11" + Chr$(72);
    End If

    For WWCicla = 1 To Lugar
    
        WPedido = Vector(WWCicla)
        Call Proceso_Click
        Call Impresion
        WMarca = "X"
            
        XParam = "'" + WPedido + "','" _
                        + WMarca + "'"
                                           
        spPedido = "ModificaPedidoImpresion1 " + XParam
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        
    
    Next WWCicla
    
    Close #1
    
    Call Cancela_click

End Sub

Private Sub Cancela_click()
    PrgIpOtro.Hide
    Unload Me
    End
End Sub

Private Sub Impresion()

    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WRazon = rstCliente!Razon
        WPago = rstCliente!Pago1
        WDirentrega = rstCliente!DirEntrega
        Rem WObservaciones = rstCliente!Observaciones
        rstCliente.Close
                
        spPago = "ConsultaPago " + "'" + WPago + "'"
        Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
        If rstPago.RecordCount > 0 Then
            WDespago = rstPago!Nombre
        End If
    End If
    

    For XX = 1 To 1

        Print #1, Tab(1); String$(79, "-")
        
        Print #1, Tab(1); "| PELLITAL S.A.";
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WPedido;
        Print #1, " / ";
        Print #1, WVersion;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Cliente";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WCliente;
        Print #1, Tab(40); Left$(WRazon, 35);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Pedido";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WFecha;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Fecha Ent.";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WFecEntrega;
        Select Case WTipoped
            Case 0
                Print #1, " (Normal)";
            Case 1
                Print #1, " (A fecha)";
            Case 2
                Print #1, " (Fecha Limite)";
            Case 3
                Print #1, " (Urgente)";
            Case Else
        End Select
            
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "C.Pago";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDespago;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Entrega";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); WDirentrega;
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); "|";
        Print #1, Tab(5); "Observaciones";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Left$(WObservaciones, 50);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|";
        Print #1, Tab(28); ":";
        Print #1, Tab(30); Right$(WObservaciones, 50);
        Print #1, Tab(80); "|"
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|";
        Print #1, Tab(2); "Producto";
        Print #1, Tab(16); "|";
        Print #1, Tab(17); "Descripcion";
        Print #1, Tab(47); "|";
        Print #1, Tab(48); "Partida";
        Print #1, Tab(58); "|";
        Print #1, Tab(59); "Cantidad";
        Print #1, Tab(67); "|";
        Print #1, Tab(68); "Envase";
        Print #1, Tab(80); "|"

        Print #1, Tab(1); String$(79, "-")
        
        XLinea = 0
        WCounter = 0
                    
        For WCounter = 1 To 100
        
            If Datos(WCounter, 0) <> "" Then
                    
                WArticulo = Datos(WCounter, 0)
                WDescripcion = Datos(WCounter, 1)
                WCantidad = Val(Datos(WCounter, 2))
                WPrecio = Val(Datos(WCounter, 3))
                WPartida = Datos(WCounter, 11)
                WObserva = Datos(WCounter, 12)
                WProceso2 = Datos(WCounter, 13)
                If Val(WProceso2) = 2 Then
                    WTarea = "Produccion"
                        Else
                    WTarea = ""
                End If
                
                
                    
                If WCantidad <> 0 Then
                    
                    Print #1, Tab(1); "|";
                    Print #1, Tab(2); WArticulo;
                    Print #1, Tab(16); "|";
                    Print #1, Tab(17); Left$(WDescripcion, 28);
                    If Trim(WPartida) <> "" Then
                        Print #1, Tab(47); "|"; Left$(WPartida, 9);
                            Else
                        Print #1, Tab(47); "|"; Left$(WTarea, 9);
                    End If
                    Print #1, Tab(58); "|";
                    Print #1, Tab(59); Alinea("#####.##", Str$(WCantidad));
                    Print #1, Tab(67); "|";
                                
                    For Cicla = 1 To 6 Step 2
                        If Val(XEnvase(WCounter, Cicla)) <> 0 Then
                            Select Case Cicla
                                Case 1
                                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvase.RecordCount > 0 Then
                                        WAbre = rstEnvase!Abreviatura
                                            Else
                                        WAbre = ""
                                    End If
                                
                                    Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                    Print #1, Tab(80); "|"
                                    
                                Case Else
                                    Print #1, Tab(1); "|";
                                    Print #1, Tab(16); "|";
                                    Print #1, Tab(47); "|";
                                    Print #1, Tab(58); "|";
                                    Print #1, Tab(67); "|";
                                                
                                    spEnvase = "ConsultaEnvases " + "'" + XEnvase(WCounter, Cicla) + "'"
                                    Set rstEnvase = db.OpenRecordset(spEnvase, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstEnvase.RecordCount > 0 Then
                                        WAbre = rstEnvase!Abreviatura
                                            Else
                                        WAbre = ""
                                    End If
                                            
                                    Print #1, Tab(68); Alinea("###", Str$(XEnvase(WCounter, Cicla + 1))) + " " + Left$(WAbre, 8);
                                    Print #1, Tab(80); "|"
                                    XLinea = XLinea + 1
                                                
                            End Select
                        End If
                    Next Cicla
                    XLinea = XLinea + 1
                    
                    If Trim(WObserva) <> 0 Then
                        
                        Print #1, Tab(1); "|";
                        Print #1, Tab(16); "|";
                        Print #1, Tab(17); Left$(WObserva, 28);
                        Print #1, Tab(47); "|";
                        Print #1, Tab(58); "|";
                        Print #1, Tab(67); "|";
                        Print #1, Tab(80); "|"
                        
                        XLinea = XLinea + 1
                        
                    End If
                    
                End If
                    
            End If
            
        Next WCounter
        
        For WDa = XLinea To 8
            Print #1, Tab(1); "|";
            Print #1, Tab(16); "|";
            Print #1, Tab(47); "|";
            Print #1, Tab(58); "|";
            Print #1, Tab(67); "|";
            Print #1, Tab(80); "|"
        Next WDa
                
     Rem agregado by nan
          
          Print #1, Tab(1); "| Preparó: ";
          Print #1, Tab(16); "|     Etiquetó";
          Print #1, Tab(47); "|";
          Print #1, Tab(58); "|";
          Print #1, Tab(67); "|";
          Print #1, Tab(80); "|"
          Print #1, Tab(1); "| Fraccionó";
          Print #1, Tab(17); "|     Supervisó";
          Print #1, Tab(47); "| Despachó";
          Print #1, Tab(58); "|";
          Print #1, Tab(67); "|";
          Print #1, Tab(80); "|"

                
                
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); "|"; WImpre(1);
        Print #1, Tab(10); "|"; WImpre(2);
        Print #1, Tab(18); "|"; WImpre(3);
        Print #1, Tab(26); "|"; WImpre(4);
        Print #1, Tab(34); "|"; WImpre(5);
        Print #1, Tab(42); "|"; WImpre(6);
        Print #1, Tab(50); "|"; WImpre(7);
        Print #1, Tab(58); "|"; WImpre(8);
        Print #1, Tab(66); "|"; WImpre(9);
        Print #1, Tab(80); "|"
        
        Print #1, Tab(1); "|(020)";
        Print #1, Tab(10); "|(021)";
        Print #1, Tab(18); "|(022)";
        Print #1, Tab(26); "|(023)";
        Print #1, Tab(34); "|(024)";
        Print #1, Tab(42); "|(025)";
        Print #1, Tab(50); "|(026)";
        Print #1, Tab(58); "|(030)";
        Print #1, Tab(66); "|(028)";
        Print #1, Tab(80); "|"
        Print #1, Tab(1); String$(79, "-")
        Print #1, Tab(1); ""
        Print #1, Tab(1); ""

    Next XX
    
    Print #1, Chr$(12)

End Sub


Private Sub Proceso_Click()

    Erase XEnvase
    Erase Datos
    Erase Auxiliar
    
    Renglon = 0
    WRenglon = 0

    spPedido = "ListaPedido " + "'" + WPedido + "'"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

    If rstPedido.RecordCount > 0 Then
            With rstPedido
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        WCliente = rstPedido!Cliente
                        WFecha = rstPedido!Fecha
                        WFecEntrega = rstPedido!FecEntrega
                        WVersion = IIf(IsNull(rstPedido!Version), "0", rstPedido!Version)
                        WTipoped = IIf(IsNull(rstPedido!Tipoped), "0", rstPedido!Tipoped)
                        WObservaciones = IIf(IsNull(rstPedido!Observaciones), "", rstPedido!Observaciones)
                        WObservaciones = Left$(WObservaciones + Space$(100), 100)
                
                        Renglon = Renglon + 1
                        
                        Datos(Renglon, 0) = rstPedido!Terminado
                        Datos(Renglon, 2) = Pusing("###,###.##", rstPedido!Cantidad - rstPedido!Facturado)
                        Datos(Renglon, 11) = IIf(IsNull(rstPedido!Partida), "", rstPedido!Partida)
                        Datos(Renglon, 12) = IIf(IsNull(rstPedido!Observa), "", rstPedido!Observa)
                        Datos(Renglon, 13) = IIf(IsNull(rstPedido!Proceso2), "", rstPedido!Proceso2)
                
                        XEnvase(Renglon, 1) = rstPedido!Envase1
                        XEnvase(Renglon, 2) = rstPedido!Canti1
                        XEnvase(Renglon, 3) = rstPedido!Envase2
                        XEnvase(Renglon, 4) = rstPedido!Canti2
                        XEnvase(Renglon, 5) = rstPedido!Envase3
                        XEnvase(Renglon, 6) = rstPedido!Canti3
                        
                        WRenglon = WRenglon + 1
                    
                        Auxiliar(WRenglon, 1) = rstPedido!Cliente
                        Auxiliar(WRenglon, 2) = rstPedido!Terminado
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPedido.Close
    End If
    
    Renglon = 0
    
    For da = 1 To WRenglon
    
        Cliente = Auxiliar(da, 1)
        Terminado = Auxiliar(da, 2)
        
        spPrecios = "ConsultaPrecios " + "'" + Cliente + Terminado + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            Renglon = Renglon + 1
            Datos(Renglon, 1) = rstPrecios!Descripcion
            Datos(Renglon, 3) = Pusing("###,###.##", rstPrecios!Precio)
        End If
        
    Next da

End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub
