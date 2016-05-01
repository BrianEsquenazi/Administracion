VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCalculaMp 
   AutoRedraw      =   -1  'True
   Caption         =   "Grabacion de Saldos de Stock y Costo de DY y DW"
   ClientHeight    =   3750
   ClientLeft      =   1515
   ClientTop       =   1245
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   3750
   ScaleWidth      =   9585
   Begin Crystal.CrystalReport Listado 
      Left            =   8160
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wStock1.rpt"
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   255
         Left            =   2400
         TabIndex        =   0
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgCalculaMp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private WVector(10000, 10) As String
Dim Empe(10, 10) As String
Dim rstStockHistorico As Recordset
Dim spStockHistorico As String
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
Dim WFechaord As String
Dim XOrden As Double
Dim XLaudo As Double
Dim XFechaOrden As String
Dim XCostoOrden As Double
Dim XParidad As Double
Dim XMoneda As Integer
Dim XTipoOrden As Integer
Dim WCosto As Double
Dim Impo1 As Double
Dim Impo2 As Double
Dim Impo3 As Double
Dim Impo4 As Double

Private Producto As String
Private Costo As Double
Private Auxiliar(100, 7) As String
Private WCodigo As String


Private Sub Acepta_Click()

    Erase WVector
    Renglon = 0

    spArticulo = "ListaArticuloStock"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo

            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Left$(rstArticulo!Codigo, 2) = "DY" Or Left$(rstArticulo!Codigo, 2) = "DW" Then
                    Renglon = Renglon + 1
                    WVector(Renglon, 1) = rstArticulo!Codigo
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    rstArticulo.Close
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaord = WAno + WMes + WDia
    
    If Val(WEmpresa) = 1 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        ZHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        ZHasta = 4
    End If
    
    
    XEmpresa = WEmpresa
    
    For Ciclo = 1 To ZHasta
            
        WEmpresa = Empe(Ciclo, 1)
        txtOdbc = Empe(Ciclo, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        For da = 1 To Renglon
    
            WEntradas = 0
            WSalidas = 0
            WArticulo = WVector(da, 1)
            XCodigo = WVector(da, 1)
            WStockActual = 0
            WStock = ""
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WStockActual = rstArticulo!Entradas - rstArticulo!Salidas
                rstArticulo.Close
                Call calcula_datos
                WStock = Str$(WStockActual - WEntradas + WSalidas)
                If Val(WStock) < 0 Then
                    WStock = "0"
                End If
            End If
            
            WVector(da, Ciclo + 1) = WStock
            WVector(da, 7) = Str$(Val(WVector(da, 7)) + Val(WStock))
            
        Next da
        
    Next Ciclo
    
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
        Case Else
    End Select
    
    For da = 1 To Renglon
    
        WArticulo = WVector(da, 1)
        XCodigo = WVector(da, 1)
            
        WStock = WVector(da, 7)
        WPlanta1 = WVector(da, 2)
        WPlanta2 = WVector(da, 3)
        WPlanta3 = WVector(da, 4)
        WPlanta4 = WVector(da, 5)
        WPlanta5 = Str$(Val(WVector(da, 6)) + Val(WVector(da, 7)) + Val(WVector(da, 8)))
        ZCosto = ""
        
        If Left$(WArticulo, 2) = "DW" Then
            WCodigo = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
            Call Calcula_Costo(WCodigo, Costo)
            ZCosto = Str$(Costo)
        End If
        
        If Val(ZCosto) = 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                ZCosto = Str$(rstArticulo!Costo2)
                rstArticulo.Close
            End If
        End If
            
        WCodigo = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
        WClave = WCodigo + Left$(WFechaord, 6)
        
        Sql1 = "DELETE StockHistorico"
        Sql2 = " Where Clave = " + "'" + WClave + "'"
        spStockHistorico = Sql1 + Sql2
        Set rstStockHistorico = db.OpenRecordset(spStockHistorico, dbOpenSnapshot, dbSQLPassThrough)
        
        Sql1 = "INSERT INTO StockHistorico ("
        Sql2 = "Clave ,"
        Sql3 = "Terminado ,"
        Sql4 = "Fecha ,"
        Sql5 = "Planta1 ,"
        Sql6 = "Planta2 ,"
        Sql7 = "Planta3 ,"
        Sql8 = "Planta4 ,"
        Sql9 = "Planta5 ,"
        Sql10 = "Costo )"
        Sql11 = "Values ("
        Sql12 = "'" + WClave + "',"
        Sql13 = "'" + WCodigo + "',"
        Sql14 = "'" + Left$(WFechaord, 6) + "',"
        Sql15 = "'" + WPlanta1 + "',"
        Sql16 = "'" + WPlanta2 + "',"
        Sql17 = "'" + WPlanta3 + "',"
        Sql18 = "'" + WPlanta4 + "',"
        Sql19 = "'" + WPlanta5 + "',"
        Sql20 = "'" + ZCosto + "')"
        
        spStockHistorico = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                           Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20
        Set rstStockHistorico = db.OpenRecordset(spStockHistorico, dbOpenSnapshot, dbSQLPassThrough)
            
    Next da
    
    Call Cancela_click
    
End Sub

Private Sub calcula_datos()

    Rem PROCESA LOS LAUDOS
    
    WEntradas = 0
    WSalidas = 0
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spLaudo = "ListaLaudoArticuloDesdeHastaFecha " + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstLaudo!Fecha, 4)
                Rem WMes = Mid$(rstLaudo!Fecha, 4, 2)
                Rem WDia = Left$(rstLaudo!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                
                Rem If WCompara > WFechaord Then
                Rem     If rstLaudo!Articulo = WArticulo Then
                        WLiberada = IIf(IsNull(rstLaudo!liberadaant), 0, rstLaudo!liberadaant)
                        If WLiberada = 0 Then
                            WLiberada = rstLaudo!Liberada
                        End If
                        WEntradas = WEntradas + WLiberada
                Rem     End If
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    Rem PROCESA LAS HOJAS DE PRODUCCION
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spHoja = "ListaHojaArticuloDesdeHastaFecha" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstHoja!Fecha, 4)
                Rem WMes = Mid$(rstHoja!Fecha, 4, 2)
                Rem WDia = Left$(rstHoja!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstHoja!Tipo = "M" And rstHoja!Articulo = WArticulo Then
                Rem         XX = rstHoja!Clave
                        Rem WCantidad = rstHoja!Canti1 + rstHoja!Canti2 + rstHoja!Canti3
                        Rem If WCantidad = 0 Then
                            WCantidad = rstHoja!Cantidad
                        Rem End If
                        WSalidas = WSalidas + WCantidad
                Rem     End If
                Rem End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstHoja!Articulo > WArticulo Then
                Rem     Exit Do
                Rem End If
                
            Loop
            End If
        
        End With
        
        rstHoja.Close
        
    End If
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spMovvar = "ListaMovvarArticuloDesdeHastaFecha" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then

        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovvar!Fecha, 4)
                Rem WMes = Mid$(rstMovvar!Fecha, 4, 2)
                Rem WDia = Left$(rstMovvar!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovvar!Tipo = "M" And rstMovvar!Articulo = WArticulo Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                    Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
                        End If
                Rem     End If
                Rem End If
                
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
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    spMovguia = "ListaMovguiaArticuloDesdeHastaFecha" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then

        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovguia!Fecha, 4)
                Rem WMes = Mid$(rstMovguia!Fecha, 4, 2)
                Rem WDia = Left$(rstMovguia!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovguia!Tipo = "M" And rstMovguia!Articulo = WArticulo Then
                        WCantidad = IIf(IsNull(rstMovguia!Cantidadant), 0, rstMovguia!Cantidadant)
                        If WCantidad = 0 Then
                            WCantidad = rstMovguia!Cantidad
                        End If
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + WCantidad
                                Else
                            WSalidas = WSalidas + WCantidad
                        End If
                Rem     End If
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
    
    
    
    Rem PROCESA LAS HOJAS DE LABORATORIO
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    
    spMovlab = "ListaMovlabArticuloDesdeHastaFecha" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem WAno = Right$(rstMovlab!Fecha, 4)
                Rem WMes = Mid$(rstMovlab!Fecha, 4, 2)
                Rem WDia = Left$(rstMovlab!Fecha, 2)
                Rem WCompara = WAno + WMes + WDia
                        
                Rem If WCompara > WFechaord Then
                Rem     If rstMovlab!Tipo = "M" And rstMovlab!Articulo = WArticulo Then
                        WCantidad = rstMovlab!Cantidad
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + WCantidad
                                Else
                            WSalidas = WSalidas + WCantidad
                        End If
                Rem     End If
                Rem End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    
    Rem PROCESA LAS VENTAS
    
    If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DW" Then
    
    XParam = "'" + WArticulo + "','" _
                 + WArticulo + "','" _
                 + WFechaord + "'"
    
    spEstadistica = "ListaEstadisticaArticuloDesdeHastaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Rem If rstEstadistica!TipoproDy = "M" And rstEstadistica!ArticuloDy = WArticulo Then
                Rem     WAno = Right$(rstEstadistica!Fecha, 4)
                Rem     WMes = Mid$(rstEstadistica!Fecha, 4, 2)
                Rem     WDia = Left$(rstEstadistica!Fecha, 2)
                Rem     WCompara = WAno + WMes + WDia
                        
                Rem     If WCompara > WFechaord Then
                        If rstEstadistica!Tipo = 1 Then
                            WSalidas = WSalidas + rstEstadistica!Cantidad
                                Else
                            WEntradas = WEntradas + rstEstadistica!Cantidad
                        End If
                Rem     End If
                Rem End If
            
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
    End If
    
    End If

End Sub

Private Sub Cancela_click()
    PrgCalculaMp.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()
    Fecha.Text = "  /  /    "
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)



    Dim Vector(100, 2) As String
    Erase Auxiliar
    ZRenglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            Entra = "S"
            
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Entra = "N"
                        
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        If Left$(Articulo1, 2) = "DW" Then
                            Tipo = "T"
                            Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    lugar = lugar + 1
                                    Vector(lugar, 1) = Articulo2
                                    Vector(lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                ZRenglon = ZRenglon + 1
                                Auxiliar(ZRenglon, 1) = Articulo1
                                Auxiliar(ZRenglon, 2) = Cantidad
                                Auxiliar(ZRenglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
            If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
                ZRenglon = ZRenglon + 1
                Auxiliar(ZRenglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
                Auxiliar(ZRenglon, 2) = 1
                Auxiliar(ZRenglon, 3) = Vector(Cicla, 2)
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For zda = 1 To ZRenglon
        Articulo = Auxiliar(zda, 1)
        Cantidad = Auxiliar(zda, 2)
        XVector = Auxiliar(zda, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = rstArticulo!Costo2
            rstArticulo.Close
        End If
        
        XOrden = 0
        XLaudo = 0
        XFechaOrden = "00000000"
        XCostoOrden = 0
        XMoneda = 0
        XTipoOrden = 0
        
        If XCostoOrden = 0 Then
            XCostoOrden = WCosto
        End If
        
        Costo = Costo + (Cantidad * XCostoOrden * Val(XVector))
        
    Next zda
    
End Sub




