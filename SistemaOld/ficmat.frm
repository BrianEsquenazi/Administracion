VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgFicmat 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ficha de Stock de Materias Primas"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wficmat.rpt"
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
      TabIndex        =   4
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
      ItemData        =   "ficmat.frx":0000
      Left            =   120
      List            =   "ficmat.frx":0007
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4680
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1200
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1680
         TabIndex        =   12
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
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
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   1440
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
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
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
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgFicmat"
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
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String
Dim Vector(10000, 7) As String
Private XLote(100, 7) As String
Private WDescripcion As String
Private WSaldo As Double
Private NombreEmpresa As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    Da = 0
    With rstFichaMat
        .Index = "Articulo"
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
    spArticulo = "ListaArticuloDesdeHasta" + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
            
        With rstArticulo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WArticulo = rstArticulo!Codigo
                WInicial = rstArticulo!Inicial
                WDescripcion = rstArticulo!Descripcion
                WFechaCierre = IIf(IsNull(rstArticulo!FechaCierre), "00/00/0000", rstArticulo!FechaCierre)
                WOrdFechaCierre = IIf(IsNull(rstArticulo!OrdFechaCierre), "00000000", rstArticulo!OrdFechaCierre)
                
                With rstFichaMat
                
                        .AddNew
                        !Articulo = WArticulo
                        !Fecha = WFechaCierre
                        !FechaOrd = "00000000"
                        !Tipo = 0
                        !Numero = 0
                        !Inicial = WInicial
                        !Entrada = 0
                        !Salida = 0
                        !Observaciones = ""
                        !Lista1 = ""
                        !Lista2 = "Saldo Inicial"
                        !Descripcion = WDescripcion
                        !Lote = 0
                        !Saldo = 0
                        .Update
                End With
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
    End If
    
    
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                If rstLaudo!Marca = "X" And rstLaudo!Saldo = 0 Then
                
                        Else
                        
                    WLiberada = IIf(IsNull(rstLaudo!Liberada), "0", rstLaudo!Liberada)
                    If WLiberada <> 0 Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!Liberada
                        WFecha = rstLaudo!Fecha
                        WLaudo = rstLaudo!Laudo
                        WOrden = rstLaudo!Orden
                        WLote = rstLaudo!Laudo
                        WSaldo = IIf(IsNull(rstLaudo!Saldo), "0", rstLaudo!Saldo)
                        Call Redondeo(WSaldo)
                                            
                        Renglon = Renglon + 1
                
                        Vector(Renglon, 1) = WArticulo
                        Vector(Renglon, 2) = WCantidad
                        Vector(Renglon, 3) = WFecha
                        Vector(Renglon, 4) = WLaudo
                        Vector(Renglon, 5) = WOrden
                        Vector(Renglon, 6) = WLote
                        Vector(Renglon, 7) = WSaldo
                        
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
    
    
    For Da = 1 To Renglon
    
        WArticulo = Vector(Da, 1)
        WCantidad = Vector(Da, 2)
        WFecha = Vector(Da, 3)
        WLaudo = Vector(Da, 4)
        WOrden = Vector(Da, 5)
        WLote = Vector(Da, 6)
        WSaldo = Vector(Da, 7)
                
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WDEsProveedor = RstProveedor!Nombre
        End If
                
        With rstFichaMat
            .AddNew
            !Articulo = WArticulo
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WLaudo
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = 0
            !Observaciones = WDEsProveedor
            !Lista1 = "Laudo"
            !Lista2 = ""
            !Lote = WLote
            !Saldo = WSaldo
            .Update
        End With
    Next Da
    
    
    
    
    Erase Vector
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                If rstLaudo!Marca = "X" Then
                
                        Else
                        
                    WDevuelta = IIf(IsNull(rstLaudo!devuelta), "0", rstLaudo!devuelta)
                    If WDevuelta <> 0 Then
                
                        WArticulo = rstLaudo!Articulo
                        WCantidad = rstLaudo!devuelta
                        WFecha = rstLaudo!Fecha
                        WLaudo = rstLaudo!Rechazo
                        WOrden = rstLaudo!Orden
                        WLote = rstLaudo!Rechazo
                        WSaldo = "0"
                    
                        Renglon = Renglon + 1
                
                        Vector(Renglon, 1) = WArticulo
                        Vector(Renglon, 2) = WCantidad
                        Vector(Renglon, 3) = WFecha
                        Vector(Renglon, 4) = WLaudo
                        Vector(Renglon, 5) = WOrden
                        Vector(Renglon, 6) = WLote
                        Vector(Renglon, 7) = WSaldo
                        
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
    
    
    For Da = 1 To Renglon
    
        WArticulo = Vector(Da, 1)
        WCantidad = Vector(Da, 2)
        WFecha = Vector(Da, 3)
        WLaudo = Vector(Da, 4)
        WOrden = Vector(Da, 5)
        WLote = Vector(Da, 6)
        WSaldo = Vector(Da, 7)
                
        spOrden = "ListaOrden" + "'" + WOrden + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
        
        WDEsProveedor = ""
                
        spProveedor = "ConsultaProveedores" + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WDEsProveedor = RstProveedor!Nombre
            RstProveedor.Close
        End If
                
        With rstFichaMat
            .AddNew
            !Articulo = WArticulo
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = WLaudo
            !Inicial = 0
            !Entrada = WCantidad
            !Salida = WCantidad
            !Observaciones = WDEsProveedor
            !Lista1 = "Rechazo"
            !Lista2 = ""
            !Lote = WLote
            Rem !WSaldo = WSaldo
            .Update
        End With
    Next Da
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                    
                        Else
                        
                    If !Tipo = "M" Then
                    
                        XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                        XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                        XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                        XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = rstHoja!Lote
                            XLote(1, 2) = rstHoja!Cantidad
                        End If
                        
                        For Da = 1 To 3
                        
                            If XLote(Da, 2) <> 0 Then
                    
                                WArticulo = rstHoja!Articulo
                                WCantidad = XLote(Da, 2)
                                WFecha = rstHoja!Fecha
                                WHoja = rstHoja!Hoja
                                WLote = XLote(Da, 1)
                                WSaldo = "0"
                
                                With rstFichaMat
                
                                    .AddNew
                                    !Articulo = WArticulo
                                    !Fecha = WFecha
                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Tipo = 0
                                    !Numero = WHoja
                                    !Inicial = 0
                                    !Entrada = 0
                                    !Salida = WCantidad
                                    !Observaciones = ""
                                    !Lista1 = "Hoja"
                                    !Lista2 = ""
                                    !Lote = WLote
                                    !Saldo = WSaldo
                                    .Update
                                End With
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                    If !Tipo = "M" Then
                
                        WArticulo = rstMovvar!Articulo
                        WCantidad = rstMovvar!Cantidad
                        WFecha = rstMovvar!Fecha
                        WCodigo = rstMovvar!Codigo
                        WMovi = rstMovvar!Movi
                        WTipomov = Val(rstMovvar!Tipomov)
                        WObservaciones = rstMovvar!Observaciones
                        WLote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                        WSaldo = "0"
                    
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
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
                            !Observaciones = WObservaciones
                            If WTipomov = 0 Or WTipomov = 1 Then
                                !Lista1 = "Mov.Var"
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
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
        rstMovvar.Close
    End If
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                    If rstMovguia!Tipo = "M" Then
                
                        WArticulo = rstMovguia!Articulo
                        WCantidad = rstMovguia!Cantidad
                        WFecha = rstMovguia!Fecha
                        WCodigo = rstMovguia!Codigo
                        WMovi = rstMovguia!Movi
                        WDestino = rstMovguia!Destino
                        WTipomov = rstMovguia!Tipomov
                        Rem WObservaciones = rstMovvar!Observaciones
                        
                        If WMovi = "E" Then
                            WLote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            WSaldo = IIf(IsNull(rstMovguia!Saldo), "0", rstMovguia!Saldo)
                            Call Redondeo(WSaldo)
                                Else
                            WLote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                            WSaldo = "0"
                        End If
                        
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
                                    WObservaciones = "Recepcion de Surfactan VI"
                                Case 11
                                    WObservaciones = "Recepcion de Surfactan VII"
                                Case Else
                                    WObservaciones = ""
                            End Select
                            
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
                                    WObservaciones = ""
                            End Select
                            
                        End If
                                
                            
                        
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
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
                            !Observaciones = WObservaciones
                            If !Numero > 900000 Then
                                !Lista1 = "Prestamo"
                                !Numero = !Numero - 900000
                                    Else
                                !Lista1 = "Guia In"
                            End If
                            !Lista2 = ""
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
    
    
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    
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
                
                If rstMovlab!Marca = "X" Then
                
                    Else
                
                    If !Tipo = "M" Then
                
                        WArticulo = rstMovlab!Articulo
                        WCantidad = rstMovlab!Cantidad
                        WFecha = rstMovlab!Fecha
                        WCodigo = rstMovlab!Codigo
                        WMovi = rstMovlab!Movi
                        WTipomov = rstMovlab!Tipomov
                        WObservaciones = rstMovlab!Observaciones
                        WLote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                        WSaldo = "0"
                        
                        With rstFichaMat
                    
                            .AddNew
                            !Articulo = WArticulo
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
                            !Observaciones = WObservaciones
                            !Lista1 = "Mov.Lab"
                            !Lista2 = ""
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
                    
                        WArticulo = rstEstadistica!ArticuloDy
                        WFecha = rstEstadistica!Fecha
                        WCodigo = rstEstadistica!Numero
                        WObservaciones = ""
                        WTipo = rstEstadistica!Tipo
                        WObservaciones = rstEstadistica!Cliente
                        
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
                        
                            WLote = XLote(Da, 1)
                            Auxi = XLote(Da, 2)
                            Auxi = Pusing("###,###.##", Auxi)
                            WCantidad = Val(Auxi)
                    
                            If WCantidad <> 0 Then
                        
                                With rstFichaMat
                    
                                    .AddNew
                                    !Articulo = WArticulo
                                    !Fecha = WFecha
                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Tipo = 0
                                    !Numero = WCodigo
                                    !Inicial = 0
                                    If WTipo = 2 Then
                                        !Entrada = Abs(WCantidad)
                                        !Salida = 0
                                        !Lista1 = "Devol."
                                            Else
                                        !Entrada = 0
                                        !Salida = WCantidad
                                        !Lista1 = "Factura"
                                    End If
                                    !Observaciones = WObservaciones
                                    !Descripcion = ""
                                    !Lista2 = ""
                                    !Lote = WLote
                                    !Saldo = 0
                                    .Update
                                End With
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
    
    
    
    
    
    
    
        
        
        
        
        
    Rem PROCESA LOS las devoluciones de mercaderia
    
    WAuxiliar = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
    
    XParam = "'" + WAuxiliar + "','" _
                 + WAuxiliar + "'"
    spEntdev = "ListaEntdevTerminadoDesdeHasta" + XParam
    Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
    If rstEntdev.RecordCount > 0 Then
    
        With rstEntdev
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEntdev!Marca = "X" Then
                
                        Else
                
                    WCantidad = rstEntdev!Cantidad
                    WFecha = rstEntdev!Fecha
                    WCodigo = rstEntdev!Codigo
                    WLote = IIf(IsNull(rstEntdev!Lote), "0", rstEntdev!Lote)
                    WPartiOri = IIf(IsNull(rstEntdev!PartiOri), "", rstEntdev!PartiOri)
                    WSaldo = rstEntdev!Saldo
                
                    With rstFichaMat
                        .AddNew
                        !Articulo = WArticulo
                        !Fecha = WFecha
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Tipo = 0
                        !Numero = WCodigo
                        !Inicial = 0
                        !Entrada = WCantidad
                        !Salida = 0
                        !Observaciones = ""
                        !Lista1 = "Ent.Dev."
                        !Lista2 = ""
                        !Lote = WLote
                        !Saldo = WSaldo
                        !PartiOri = WPartiOri
                        .Update
                    End With
                
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
    
   
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Da = 0
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                WArticulo = !Articulo
                WObservaciones = !Observaciones
                WDescripcion = ""
                WNumero = !Lote
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescripcion = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                If !Lista1 = "Devol." Or !Lista1 = "Factura" Then
                    spCliente = "ConsultaCliente" + "'" + WObservaciones + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WObservaciones = rstCliente!Razon
                        rstCliente.Close
                    End If
                End If
                !Descripcion = WDescripcion
                !Observaciones = WObservaciones
                
                Select Case Left$(WArticulo, 2)
                    Case "DY", "DS", "DQ"
                        WPartiOri = ""
                        If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DS" Or Left$(WArticulo, 2) = "DQ" Then
                            spLaudo = "ListaLaudo " + "'" + WNumero + "'"
                            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                            If rstLaudo.RecordCount > 0 Then
                                WPartiOri = IIf(IsNull(rstLaudo!PartiOri), "", rstLaudo!PartiOri)
                                rstLaudo.Close
                            End If
                        End If
                        !PartiOri = WPartiOri
                    Case "DK", "NS"
                        Rem no hago nada
                    Case Else
                        !PartiOri = ""
                End Select
                
                !Empresa = NombreEmpresa
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    

    Listado.WindowTitle = "Listado de Ficha de Stock de Materias Primas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{FichaMat.Articulo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstFichaMat
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgFicmat.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaMat
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgFicmat.Caption = "Listado de Ficha de Stock de Materias Primas :  " + !Nombre
            NombreEmpresa = !Nombre
        End If
    End With
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstArticulo!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WArticulo = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
            Desde.Text = rstArticulo!Codigo
            Hasta.Text = rstArticulo!Codigo
                Else
            Desde.Text = WArticulo
            Hasta.Text = WArticulo
    End If
    Desde.SetFocus
    
    
End Sub

