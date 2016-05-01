VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgLoteter1 
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
      ItemData        =   "Loteter1.frx":0000
      Left            =   120
      List            =   "Loteter1.frx":0007
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
Attribute VB_Name = "PrgLoteter1"
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


Private Sub Acepta_Click()

    Lote.Text = Pasalote

    Erase Vector
    Renglon = 0

    da = 0
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
        WFecha = rstHoja!Fecha
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
                
                Rem If rstEstadistica!numero = 9473 Then Stop
                
                If rstEstadistica!lote1 = Val(Lote.Text) Then
                
                    WTipo = rstEstadistica!Tipo
                    WTerminado = rstEstadistica!Articulo
                    WSalida = rstEstadistica!Canti1
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
                
                If rstEstadistica!lote2 = Val(Lote.Text) Then
                
                    WTipo = rstEstadistica!Tipo
                    WTerminado = rstEstadistica!Articulo
                    WSalida = rstEstadistica!Canti2
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
                
                If rstEstadistica!lote3 = Val(Lote.Text) Then
                
                    WTipo = rstEstadistica!Tipo
                    WTerminado = rstEstadistica!Articulo
                    WSalida = rstEstadistica!Canti3
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
                
                If rstEstadistica!lote4 = Val(Lote.Text) Then
                
                    WTipo = rstEstadistica!Tipo
                    WTerminado = rstEstadistica!Articulo
                    WSalida = rstEstadistica!Canti4
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
                
                If rstEstadistica!lote5 = Val(Lote.Text) Then
                
                    WTipo = rstEstadistica!Tipo
                    WTerminado = rstEstadistica!Articulo
                    WSalida = rstEstadistica!Canti5
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
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstEstadistica.Close
    End If
    
    For da = 1 To Renglon
    
        WTipo = Val(Vector(da, 1))
        WTerminado = Vector(da, 2)
        WSalida = Val(Vector(da, 3))
        WFecha = Vector(da, 4)
        WNumero = Vector(da, 5)
        WImpre1 = Vector(da, 6)
        
        spCliente = "ConsultaCliente" + "'" + WImpre1 + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WImpre2 = rstCliente!Razon
                Else
            WImpre2 = ""
        End If
                
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
    Next da
    
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
                
                If rstHoja!Marca = "X" Then
                
                    Else
                
                If rstHoja!Tipo = "T" Then
                
                    WTerminado = rstHoja!Terminado
                    WCantidad = rstHoja!Cantidad
                    WFecha = rstHoja!Fecha
                    WHoja = rstHoja!Hoja
                    
                    If rstHoja!lote1 = Val(Lote.Text) Then
                
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
                                    WObservacionesWObservaciones = "Envio a Pelitall"
                                Case 3
                                    WObservaciones = "Envio a Surfactan II"
                                Case 4
                                    WObservaciones = "Envio a Pelitall II"
                                Case 5
                                    WObservaciones = "Envio a Surfactan III"
                                Case 6
                                    WObservaciones = "Envio a Surfactan IV"
                                Case Else
                            End Select
                            WLote = rstMovguia!Partida
                            WSaldo = 0
                            
                                Else
                                
                            Select Case WTipomov
                                Case 1
                                    WObservaciones = "Recepcion de Surfactan"
                                Case 2
                                    WObservaciones = "Recepcion de Pelitall"
                                Case 3
                                    WObservaciones = "Recepcion de Surfactan II"
                                Case 4
                                    WObservaciones = "Recepcion de Pelitall II"
                                Case 5
                                    WObservaciones = "Recepcion de Surfactan III"
                                Case 6
                                    WObservaciones = "Recepcion de Surfactan IV"
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
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        
        End With
        rstMovlab.Close
    End If
    
    da = 0
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
                    spCliente = "ConsultaCliente " + "'" + Left$(!Observaciones, 6) + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        !Lista2 = Left$(rstCliente!Razon, 30)
                        rstCliente.Close
                    End If
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
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
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
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgConsFicTer.Show
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLoteter.Caption = "Listado de Ficha de Lote de Productos Terminados :  " + !Nombre
        End If
    End With
    Lote.Text = Pasalote
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = True
    Call Acepta_Click
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
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


