VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgValuaMatPartida 
   AutoRedraw      =   -1  'True
   Caption         =   "Calculo de Precio Promedio"
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
      ReportFileName  =   "WValuaMatPartida.rpt"
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
      ItemData        =   "ValuaMatPartida.frx":0000
      Left            =   120
      List            =   "ValuaMatPartida.frx":0007
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
         TabIndex        =   10
         Top             =   1080
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
         Top             =   1080
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
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgValuaMatPartida"
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
Dim ZCosto As Double


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
                                            
                        If WSaldo > 0 Then
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = WArticulo
                            Vector(Renglon, 2) = Str$(WCantidad)
                            Vector(Renglon, 3) = WFecha
                            Vector(Renglon, 4) = Str$(WLaudo)
                            Vector(Renglon, 5) = WOrden
                            Vector(Renglon, 6) = Str$(WLote)
                            Vector(Renglon, 7) = Str$(WSaldo)
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
                                    WObservaciones = "Envio a Surfactan VI"
                                Case 11
                                    WObservaciones = "Envio a Surfactan VII"
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
                                
                            
                        If WSaldo > 0 Then
                        
                            Renglon = Renglon + 1
                            Vector(Renglon, 1) = WArticulo
                            Vector(Renglon, 2) = WCantidad
                            Vector(Renglon, 3) = WFecha
                            Vector(Renglon, 4) = WLote
                            Vector(Renglon, 5) = ""
                            Vector(Renglon, 6) = WLote
                            Vector(Renglon, 7) = WSaldo
                            
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
    
    
    
    
    
    
    
    
    
    
    
    
    
    For Da = 1 To Renglon
    
        WArticulo = Vector(Da, 1)
        WCantidad = Val(Vector(Da, 2))
        WFecha = Vector(Da, 3)
        WLaudo = Vector(Da, 4)
        WOrden = Vector(Da, 5)
        WLote = Vector(Da, 6)
        WSaldo = Val(Vector(Da, 7))
        
        WCosto = 0
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Articulo.Codigo = " + "'" + WArticulo + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = rstArticulo!Costo1
            WDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CostoPartida"
        ZSql = ZSql + " Where CostoPartida.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and CostoPartida.Laudo = " + "'" + WLaudo + "'"
        spCostoPartida = ZSql
        Set rstCostoPartida = db.OpenRecordset(spCostoPartida, dbOpenSnapshot, dbSQLPassThrough)
        If rstCostoPartida.RecordCount > 0 Then
            WCosto = rstCostoPartida!Costo
            rstCostoPartida.Close
        End If
                
        With rstFichaMat
            .AddNew
            !Articulo = WArticulo
            !Descripcion = WDescripcion
            !Fecha = WFecha
            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
            !Tipo = 0
            !Numero = Val(WLaudo)
            !Inicial = 0
            !Entrada = WSaldo
            !Salida = WCosto
            !Observaciones = ""
            !Lista1 = ""
            !Lista2 = ""
            !Lote = WLote
            !Saldo = WSaldo
            .Update
        End With
        
    Next Da
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Articulo SET "
    ZSql = ZSql + " Costo3 = Costo1"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    ZPasa = 0
    ZCorte = ""
    ZSumaImpo = 0
    ZSumaCanti = 0
    
    With rstFichaMat
        .Index = "Articulo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                If ZPasa = 0 Then
                    ZPasa = 1
                    ZCorte = !Articulo
                End If
                
                If ZCorte <> !Articulo Then
                
                    If ZSumaCanti <> 0 And ZSumaImpo <> 0 Then
                        ZCosto = ZSumaImpo / ZSumaCanti
                        Call Redondeo(ZCosto)
                        
                    
                        ZSql = ""
                        ZSql = ZSql + "UPDATE Articulo SET "
                        ZSql = ZSql + " Costo3 = " + "'" + Str$(ZCosto) + "'"
                        ZSql = ZSql + " Where Codigo = " + "'" + ZCorte + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                
                    ZCorte = !Articulo
                    ZSumaImpo = 0
                    ZSumaCanti = 0
                    
                End If
                Impo = !Entrada * !Salida
                ZSumaImpo = ZSumaImpo + Impo
                ZSumaCanti = ZSumaCanti + !Entrada
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    If ZPasa <> 0 Then
        If ZSumaCanti <> 0 And ZSumaImpo <> 0 Then
            ZCosto = ZSumaImpo / ZSumaCanti
            Call Redondeo(ZCosto)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Articulo SET "
            ZSql = ZSql + " Costo3 = " + "'" + Str$(ZCosto) + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + ZCorte + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        End If
    End If
    
    Rem Desde.Text = UCase(Desde.Text)
    Rem Hasta.Text = UCase(Hasta.Text)
    
    Rem Listado.GroupSelectionFormula = "{@Stock} > 0 and {Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    Rem Listado.SelectionFormula = "{Articulo.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    Rem If Impresora.Value = True Then
    Rem     Listado.Destination = 1
    Rem         Else
    Rem     Listado.Destination = 0
    Rem End If
    
    Rem DbConnect = db.Connect
    Rem DSQ = getDatabase(DbConnect)
    
    Rem Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Costo3, Articulo.Inicial, Articulo.Entradas, Articulo.Salidas " _
    rem             + "From " _
    rem             + DSQ + ".dbo.Articulo Articulo " _
    rem             + "Where " _
    rem             + "Articulo.Codigo >= '" + Desde.Text + "' AND " _
    rem             + "Articulo.Codigo <= '" + Hasta.Text + "'"
    Rem
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Rem Listado.Connect = Connect()
    
    Rem Listado.Action = 1
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstFichaMat
        .Close
    End With
    
    DbsEmpresa.Close
    
    PrgValuaMatPartida.Hide
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

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Desde.Text = "AA-000-000"
    Hasta.Text = "ZZ-999-999"
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

