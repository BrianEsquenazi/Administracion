VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLoteter 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ficha de Lote de Producto Terminado"
   ClientHeight    =   6180
   ClientLeft      =   2085
   ClientTop       =   1500
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Lote 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Lote"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5760
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wloteter.rpt"
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
      Left            =   6240
      TabIndex        =   3
      Top             =   120
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
      Height          =   3960
      ItemData        =   "loteter.frx":0000
      Left            =   120
      List            =   "loteter.frx":0007
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "PrgLoteter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String

Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovguia As String
Dim rstMovguia As Recordset
Dim spMovvar As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String

Dim XParam As String
Dim Vector(10000, 6) As String
Private XLote(100, 7) As String

Private Sub Acepta_Click()

    On Error GoTo WError

    Erase Vector
    Renglon = 0

    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
            
    XParam = "'" + Lote.Text + "'"
    spHoja = "ListaHoja" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        WProducto = rstHoja!Producto
        WCantidad = rstHoja!Real
        WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
        WFechaFinal = Trim(WFechaFinal)
        If WFechaFinal <> "" Then
            WFecha = WFechaFinal
                Else
            WFecha = rstHoja!Fecha
        End If
        Rem WFecha = rstHoja!Fecha
        WHoja = rstHoja!Hoja
        WSaldo = rstHoja!Saldo
                
        With rstFichaTer
                
            .AddNew
            !Terminado = WProducto
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WHoja
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = 0
            !Observaciones = ""
            !Lista1 = "Hoja"
            !Lista2 = ""
            !Lote = WHoja
            !Saldo = WSaldo
            .Update
        End With
        
        rstHoja.Close
        
            Else
            
        XParam = "'" + Lote.Text + "'"
        spMovguia = "ListaMovguiaLoteSolo" + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
            WProducto = rstMovguia!Terminado
            rstMovguia.Close
        End If
        
    End If
    
    spTerminado = "ConsultaTerminado " + "'" + WProducto + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
        rstTerminado.Close
    End If
    
            
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                If rstEstadistica!Marca = "X" Then
                
                        Else
                        
                    XLote(1, 1) = IIf(IsNull(rstEstadistica!Lote1), "", rstEstadistica!Lote1)
                    XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstEstadistica!Lote2), "", rstEstadistica!Lote2)
                    XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstEstadistica!Lote3), "", rstEstadistica!Lote3)
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
            
                        WLote = XLote(Da, 1)
                        WCanti = XLote(Da, 2)
                        
                        If Val(WLote) = Val(Lote.Text) Then
                        
                            WTipo = rstEstadistica!Tipo
                            WTerminado = rstEstadistica!Articulo
                            WSalida = Val(WCanti)
                            WFecha = rstEstadistica!Fecha
                            WNumero = rstEstadistica!Numero
                            WImpre1 = rstEstadistica!Cliente
                            
                            Renglon = Renglon + 1
                        
                            Vector(Renglon, 1) = WTipo
                            Vector(Renglon, 2) = WTerminado
                            Vector(Renglon, 3) = WSalida
                            Vector(Renglon, 4) = WFecha
                            Vector(Renglon, 5) = WNumero
                            Vector(Renglon, 6) = WImpre1
                        
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
    
    For Da = 1 To Renglon
    
        WTipo = Val(Vector(Da, 1))
        WTerminado = Vector(Da, 2)
        WSalida = Val(Vector(Da, 3))
        WFecha = Vector(Da, 4)
        WNumero = Vector(Da, 5)
        WImpre1 = Vector(Da, 6)
        
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WImpre2 = rstCliente!Razon
            rstCliente.Close
                Else
            WImpre2 = ""
        End If
        
        Call Conecta_Empresa
                
        With rstFichaTer
                
                .AddNew
                !Terminado = WTerminado
                !Fecha = WFecha
                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                !Tipo = 0
                !Numero = WNumero
                !Inicial = 0
                If Val(WTipo) = 1 Then
                    !Entrada = 0
                    !Salida = WSalida
                    !Lista1 = "Facura"
                            Else
                    !Salida = 0
                    !Entrada = Abs(WSalida)
                    !Lista1 = "Devol"
                End If
                !Observaciones = ""
                !Lista2 = WImpre1 + " " + Left$(WImpre2, 23)
                !Lote = Lote.Text
                !Saldo = 0
                .Update
        End With
    Next Da
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                    WTerminado = rstHoja!Terminado
                    WCantidad = rstHoja!Cantidad
                    WFechaFinal = IIf(IsNull(rstHoja!FechaFinal), "", rstHoja!FechaFinal)
                    WFechaFinal = Trim(WFechaFinal)
                    If WFechaFinal <> "" Then
                        WFecha = WFechaFinal
                            Else
                        WFecha = rstHoja!Fecha
                    End If
                    Rem WFecha = rstHoja!Fecha
                    WHoja = rstHoja!Hoja
                    
                    If rstHoja!Lote1 = Val(Lote.Text) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = rstHoja!Canti1
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Lote.Text
                            !Saldo = 0
                            .Update
                        End With
                        
                    End If
                    
                    If rstHoja!Lote2 = Val(Lote.Text) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = rstHoja!Canti2
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Lote.Text
                            !Saldo = 0
                            .Update
                        End With
                        
                    End If
                    
                    If rstHoja!Lote3 = Val(Lote.Text) Then
                
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WHoja
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = rstHoja!Canti3
                            !Observaciones = ""
                            !Lista1 = "Hoja"
                            !Lista2 = ""
                            !Lote = Lote.Text
                            !Saldo = 0
                            .Update
                        End With
                        
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
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                If rstMovvar!Tipo = "T" Then
                
                    WTerminado = rstMovvar!Terminado
                    WCantidad = rstMovvar!Cantidad
                    WFecha = rstMovvar!Fecha
                    WCodigo = rstMovvar!Codigo
                    WMovi = rstMovvar!Movi
                    WTipomov = Val(rstMovvar!Tipomov)
                    WObservaciones = rstMovvar!Observaciones
                    WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                    
                    If Val(WLote) = Val(Lote.Text) Then

                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            If WTipomov = 1 Or WTipomov = 2 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = 0
                            .Update
                        End With
                        
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
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                If rstMovguia!Marca = "X" Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    WTerminado = rstMovguia!Terminado
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    Rem WObservaciones = rstMovvar!Observaciones
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    
                    If WMovi = "S" Then
                            Select Case WDestino
                                Case 1
                                    WObservaciones = "Envio a Surfactan"
                                Case 2
                                    WObservaciones = "Envio a Pellital"
                                Case 3
                                    WObservaciones = "Envio a Surfactan II"
                                Case 4
                                    WObservaciones = "Envio a Pellital II"
                                Case 5
                                    WObservaciones = "Envio a Surfactan III"
                                Case 6
                                    WObservaciones = "Envio a Surfactan IV"
                                Case 7
                                    WObservaciones = "Envio a Surfactan V"
                                Case 8
                                    WObservaciones = "Envio a Pellital V"
                                Case 9
                                    WObservaciones = "Envio a Pellital IV"
                                Case 10
                                    WObservaciones = "Envio a Surfactan VI"
                                Case 11
                                    WObservaciones = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Partida
                            WSaldo = 0
                            
                                Else
                                
                            Select Case WTipomov
                                Case 1
                                    WObservaciones = "Recepcion de Surfactan"
                                Case 2
                                    WObservaciones = "Recepcion de Pellital"
                                Case 3
                                    WObservaciones = "Recepcion de Surfactan II"
                                Case 4
                                    WObservaciones = "Recepcion de Pellital II"
                                Case 5
                                    WObservaciones = "Recepcion de Surfactan III"
                                Case 6
                                    WObservaciones = "Recepcion de Surfactan IV"
                                Case 7
                                    WObservaciones = "Recepcion de Surfactan V"
                                Case 8
                                    WObservaciones = "Recepcion de Pellital V"
                                Case 9
                                    WObservaciones = "Recepcion de Pellital IV"
                                Case 10
                                    WObservaciones = "Recepcion de Surfactan VI"
                                Case 11
                                    WObservaciones = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Lote
                            WSaldo = rstMovguia!Saldo
                            
                    End If
                    
                    If WLote = Val(Lote.Text) Then
                        
                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            !Lista1 = "Guia In"
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = WSaldo
                            .Update
                        End With
                    
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
    
    
    
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                    WTerminado = rstConsig!Terminado
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WFecha = rstConsig!Fecha
                    WCodigo = rstConsig!Numero
                    WCliente = rstConsig!Cliente
                    WObservaciones = rstConsig!Observaciones
                    
                    If WCantidad <> 0 Then

                        With rstFichaTer
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = WCantidad
                            !Observaciones = WCliente
                            !Lista1 = "Rem.Con."
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = Lote.Text
                            !Saldo = 0
                            .Update
                        End With
                        
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
    
    XParam = "'" + WProducto + "','" _
                 + WProducto + "'"
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
                
                If rstMovlab!Marca = "X" Then
                
                        Else
                
                If rstMovlab!Tipo = "T" Then
                
                    WTerminado = rstMovlab!Terminado
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = rstMovlab!Lote
                    
                    If WLote = Val(Lote.Text) Then

                        With rstFichaTer
                
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            If WMovi = "E" Then
                                !Entrada = WCantidad
                                !Salida = 0
                                    Else
                                !Entrada = 0
                                !Salida = WCantidad
                            End If
                            !Observaciones = ""
                            !Lista1 = "Mov.Lab"
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = 0
                            .Update
                        End With
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
    
    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WTerminado = !Terminado
                WDescripcion = ""
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                !Descripcion = WDescripcion
                
                If Left$(!Lista1, 8) = "Rem.Con." Then
                
                    WLista2 = ""
                
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0008"
                            txtOdbc = "Empresa08"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
                
                    spCliente = "ConsultaCliente " + "'" + Left$(!Observaciones, 6) + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WLista2 = Left$(rstCliente!Razon, 30)
                        rstCliente.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    !Lista2 = WLista2
                    
                End If
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Ficha de Lote de Productos Terminados"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaTer.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub

WError:

    Resume Next
    
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    With rstFichaTer
        .Close
    End With
    DbsEmpresa.Close
    
    Lote.SetFocus
    PrgLoteter.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaTer
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLoteter.Caption = "Listado de Ficha de Lote de Productos Terminados :  " + !Nombre
        End If
    End With
    Lote.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstTerminado!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Desde.Text = rstTerminado!Codigo
        Hasta.Text = rstTerminado!Codigo
        rstTerminado.Close
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub




