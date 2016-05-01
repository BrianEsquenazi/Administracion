VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgStock2Otro 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valorizacion de Producto Terminado a Fecha"
   ClientHeight    =   4125
   ClientLeft      =   210
   ClientTop       =   1410
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   4125
   ScaleWidth      =   11655
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WStock2.rpt"
   End
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   5415
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
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
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
         Height          =   375
         Left            =   960
         TabIndex        =   4
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3840
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Producto"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
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
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgStock2Otro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Private WEntradas As Double
Private WSalidas As Double
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
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
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Dim WFechaord As String
Dim Impo1 As Double
Dim Impo2 As Double
Dim Impo3 As Double
Dim Impo4 As Double
Private Producto As String
Private Costo As Double
Private Auxiliar(180, 7) As String
Private WVector(20000, 2) As String
Dim Empe(12, 10) As String
Private WCodigo As String

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Fecha.SetFocus
    PrgStock2Otro.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Acepta_Click()

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Posdat = "al " + Fecha.Text
            .Update
        End If
    End With


    Erase WVector
    Renglon = 0
        
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
        .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstTerminado!Codigo >= Desde.Text And rstTerminado!Codigo <= Hasta.Text Then
                          
                    Renglon = Renglon + 1
                    WVector(Renglon, 1) = rstTerminado!Codigo
                    WStock = Str$(rstTerminado!Entradas - rstTerminado!Salidas)
                    WVector(Renglon, 2) = WStock
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With
    rstTerminado.Close
    
    End If
    
    WAno = Right$(Fecha.Text, 4)
    WMes = Mid$(Fecha.Text, 4, 2)
    WDia = Left$(Fecha.Text, 2)
    WFechaord = WAno + WMes + WDia
    
    spTerminado = "ModificaTerminadoStock0"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WTerminado = WVector(Da, 1)
        XCodigo = WVector(Da, 1)
        WCodigo = WVector(Da, 1)
        XStock = Val(WVector(Da, 2))
        XDate = Date$
        
        Call calcula_datos
        
        WStock = Str$(XStock - WEntradas + WSalidas)
        Impo1 = XStock
        Impo2 = WEntradas
        Impo3 = WSalidas
        Impo4 = Impo1 - Impo2 + Impo3
        
        Call Redondeo(Impo1)
        Call Redondeo(Impo2)
        Call Redondeo(Impo3)
        Call Redondeo(Impo4)
        
        If Impo4 < 0 Then
            Impo4 = 0
        End If
        
        WStock = Str$(Impo4)
        Costo = 0
        
        If Val(WStock) > 0 Then
            Call Calcula_Costo(WCodigo, Costo)
        End If
        WCosto = Str$(Costo)
        
        XParam = "'" + XCodigo + "','" _
                + WStock + "','" _
                + WCosto + "'"
                                           
        spTerminado = "ModificaTerminadoStock " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da
    
    Listado.WindowTitle = "Listado de Valorizacion de Producto Terminado a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaEnv.Envase} in " + DesdeEnv.Text + " to " + HastaEnv.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.Linea, Terminado.Costo, Terminado.Stock " _
                        + "From " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "Terminado.Codigo >= '  -     -   ' AND Terminado.Codigo <= 'ZZ-99999-999' AND Terminado.Stock <> 0."
    
    Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1

End Sub

Private Sub calcula_datos()

    WEntradas = 0
    WSalidas = 0

    Rem PROCESA LAS ESTADISTICAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
                 
    spEstadistica = "ListaEstadisticaDesdeHastaFecha" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstEstadistica!Fecha, 4)
                WMes = Mid$(rstEstadistica!Fecha, 4, 2)
                WDia = Left$(rstEstadistica!Fecha, 2)
                WCompara = WAno + WMes + WDia
           Rem BY NAN
                
                
                
                
                
                
                If WCompara > WFechaord Then
                       Rem BY NAN
                    
                     If WEmpresa = "0001" Then
                         lin = rstEstadistica!Linea
                            If lin = 6 Or lin = 8 Or lin = 16 Then
                       
                
                              If WCompara > WFechaord Then
                                 If Val(rstEstadistica!Tipo) = 1 Then
                                    WSalidas = WSalidas + rstEstadistica!Cantidad
                                       Else
                                   WEntradas = WEntradas + Abs(rstEstadistica!Cantidad)
                                  End If
                                 End If
                                  
                             End If
                
                Else
                    
                    
                                
                    
                    
                    
                    
                    
                    
                    If Val(rstEstadistica!Tipo) = 1 Then
                        WSalidas = WSalidas + rstEstadistica!Cantidad
                            Else
                        WEntradas = WEntradas + Abs(rstEstadistica!Cantidad)
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
    
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
    spHoja = "ListaHojaDesdeHastaFecha" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstHoja!Fecha, 4)
                WMes = Mid$(rstHoja!Fecha, 4, 2)
                WDia = Left$(rstHoja!Fecha, 2)
                WCompara = WAno + WMes + WDia
                
                If WCompara > WFechaord Then
                    If rstHoja!Tipo = "T" Then
                        WSalidas = WSalidas + rstHoja!Cantidad
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
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
    spHoja = "ListaHojaProductoDesdeHastaFecha" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                
                WAno = Right$(rstHoja!Fecha, 4)
                WMes = Mid$(rstHoja!Fecha, 4, 2)
                WDia = Left$(rstHoja!Fecha, 2)
                WCompara = WAno + WMes + WDia
                
                
                If WCompara > WFechaord Then
                    WCantidad = IIf(IsNull(rstHoja!realant), 0, rstHoja!realant)
                    If WCantidad = 0 Then
                        WCantidad = rstHoja!Real
                    End If
                    If Val(rstHoja!Renglon) = 1 And WCantidad <> 0 Then
                         WEntradas = WEntradas + WCantidad
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
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
    spMovvar = "ListaMovvarTerminadoDesdeHastaFecha" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstMovvar!Fecha, 4)
                WMes = Mid$(rstMovvar!Fecha, 4, 2)
                WDia = Left$(rstMovvar!Fecha, 2)
                WCompara = WAno + WMes + WDia
                
                If WCompara > WFechaord Then
                    If rstMovvar!Tipo = "T" Then
                        If rstMovvar!Movi = "E" Then
                            WEntradas = WEntradas + rstMovvar!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovvar!Cantidad
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
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
    spMovguia = "ListaMovguiaTerminadoDesdeHastaFecha" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstMovguia!Fecha, 4)
                WMes = Mid$(rstMovguia!Fecha, 4, 2)
                WDia = Left$(rstMovguia!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara > WFechaord Then
                    If rstMovguia!Tipo = "T" Then
                        WCantidad = IIf(IsNull(rstMovguia!Cantidadant), 0, rstMovguia!Cantidadant)
                        If WCantidad = 0 Then
                            WCantidad = rstMovguia!Cantidad
                        End If
                        If rstMovguia!Movi = "E" Then
                            WEntradas = WEntradas + WCantidad
                                Else
                            WSalidas = WSalidas + WCantidad
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
    
    
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "','" _
                 + WFechaord + "'"
    spMovlab = "ListaMovlabTerminadoDesdeHastaFecha" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstMovlab!Fecha, 4)
                WMes = Mid$(rstMovlab!Fecha, 4, 2)
                WDia = Left$(rstMovlab!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara > WFechaord Then
                    If rstMovlab!Tipo = "T" Then
                        If rstMovlab!Movi = "E" Then
                            WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                            WSalidas = WSalidas + rstMovlab!Cantidad
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
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spConsig = "ListaConsigRepro" + XParam
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
    
        With rstConsig
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                WAno = Right$(rstConsig!Fecha, 4)
                WMes = Mid$(rstConsig!Fecha, 4, 2)
                WDia = Left$(rstConsig!Fecha, 2)
                WCompara = WAno + WMes + WDia
                    
                If WCompara > WFechaord Then
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WSalidas = WSalidas + WCantidad
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
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Fecha.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgStock2Otro.Caption = "Listado de Valorizacion de Producto Terminado a Fecha :  " + !Nombre
        End If
    End With
    
    Fecha.Text = "  /  /    "
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
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
                        
                        Rem If Left$(Articulo1, 2) = "DW" Then
                        Rem     Tipo = "T"
                        Rem     Articulo2 = Left$(Articulo1, 3) + "00" + Right$(Articulo1, 7)
                        Rem End If
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
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
            
            Rem If Entra = "S" And Left$(Vector(Cicla, 1), 2) = "DW" Then
            Rem     Renglon = Renglon + 1
            Rem     Auxiliar(Renglon, 1) = Left$(Vector(Cicla, 1), 3) + Right$(Vector(Cicla, 1), 7)
            Rem     Auxiliar(Renglon, 2) = 1
            Rem     Auxiliar(Renglon, 3) = Vector(Cicla, 2)
            Rem End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        XVector = Auxiliar(Da, 3)
        
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
        
        Rem XEmpresa = WEmpresa
        Rem
        Rem If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 9  Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Rem     Empe(1, 1) = "0001"
        Rem     Empe(1, 2) = "Empresa01"
        Rem     Empe(2, 1) = "0003"
        Rem     Empe(2, 2) = "Empresa03"
        Rem     Empe(3, 1) = "0005"
        Rem     Empe(3, 2) = "Empresa05"
        Rem     Empe(4, 1) = "0006"
        Rem     Empe(4, 2) = "Empresa06"
        Rem     Empe(5, 1) = "0007"
        Rem     Empe(5, 2) = "Empresa07"
        Rem     XHasta = 5
        Rem         Else
        Rem     Empe(1, 1) = "0002"
        Rem     Empe(1, 2) = "Empresa02"
        Rem     Empe(2, 1) = "0004"
        Rem     Empe(2, 2) = "Empresa04"
        Rem     Empe(3, 1) = "0008"
        Rem     Empe(3, 2) = "Empresa08"
        Rem     XHasta = 3
        Rem End If
        Rem
        Rem For A = 1 To XHasta
        Rem
        Rem     WEmpresa = Empe(A, 1)
        Rem     txtOdbc = Empe(A, 2)
        Rem     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem
        Rem     XOrden = 0
        Rem     XCodigo = Articulo
        Rem
        Rem     XParam = "'" + XCodigo + "','" _
        REM          + XCodigo + "'"
        Rem     spLaudo = "ListaLaudoArticuloDesdeHasta" + XParam
        Rem     Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        Rem     If rstLaudo.RecordCount > 0 Then
        Rem         With rstLaudo
        Rem             .MoveFirst
        Rem             If .NoMatch = False Then
        Rem             Do
        Rem                 If .EOF = True Then
        Rem                     Exit Do
        Rem                 End If
        Rem                 If rstLaudo!Articulo = XCodigo Then
        Rem                     XOrdFecha = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
        Rem                     If XOrdFecha <= WFechaord Then
        Rem                     If XOrdFecha > XFechaOrden Then
        Rem                         XFechaOrden = XOrdFecha
        Rem                         XOrden = !Orden
        Rem                         XLaudo = !Laudo
        Rem                     End If
        Rem                     End If
        Rem                 End If
        Rem                 .MoveNext
        Rem                 If .EOF = True Then
        Rem                     Exit Do
        Rem                 End If
        Rem             Loop
        Rem             End If
        Rem         End With
        Rem         rstLaudo.Close
        Rem     End If
        Rem
        Rem     If XOrden <> 0 Then
        Rem         spOrden = "ListaOrdenArticulo " + "'" + Str$(XOrden) + "','" + XCodigo + "'"
        Rem         Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstOrden.RecordCount > 0 Then
        Rem             XTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
        Rem             Rem If XTipoOrden = 0 And rstOrden!Precio <> 1 Then
        Rem             If rstOrden!Precio <> 1 Then
        Rem                 XCostoOrden = rstOrden!Precio
        Rem                 XMoneda = IIf(IsNull(rstOrden!Moneda), "0", rstOrden!Moneda)
        Rem                     Else
        Rem                 vercodigo = XCodigo
        Rem                 XFechaOrden = WFechaord
        Rem                 XCostoOrden = WCosto
        Rem                 XMoneda = 0
        Rem             End If
        Rem             rstOrden.Close
        Rem         End If
        Rem         If XMoneda = 0 Then
        Rem             WEmpresa = "0001"
        Rem             txtOdbc = "Empresa01"
        Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem
        Rem             spCambios = "ConsultaCambioOrdFecha  " + "'" + XFechaOrden + "'"
        Rem             Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        Rem             If rstCambios.RecordCount > 0 Then
        Rem                 With rstCambios
        Rem                     .MoveLast
        Rem                     XParidad = rstCambios!Cambio
        Rem                     rstCambios.Close
        Rem                 End With
        Rem                     Else
        Rem                 XParidad = 1
        Rem             End If
        Rem
        Rem             WEmpresa = Empe(A, 1)
        Rem             txtOdbc = Empe(A, 2)
        Rem             strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem             XCostoOrden = XCostoOrden * XParidad
        Rem         End If
        Rem     End If
        Rem
        Rem Next A
        Rem
        Rem Select Case Val(XEmpresa)
        Rem     Case 1
        Rem         WEmpresa = "0001"
        Rem         txtOdbc = "Empresa01"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 2
        Rem         WEmpresa = "0002"
        Rem         txtOdbc = "Empresa02"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 3
        Rem         WEmpresa = "0003"
        Rem         txtOdbc = "Empresa03"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 4
        Rem         WEmpresa = "0004"
        Rem         txtOdbc = "Empresa04"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 5
        Rem         WEmpresa = "0005"
        Rem         txtOdbc = "Empresa05"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 6
        Rem         WEmpresa = "0006"
        Rem         txtOdbc = "Empresa06"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 7
        Rem         WEmpresa = "0007"
        Rem         txtOdbc = "Empresa07"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 8
        Rem         WEmpresa = "0008"
        Rem         txtOdbc = "Empresa08"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case 9
        Rem         WEmpresa = "0009"
        Rem         txtOdbc = "Empresa09"
        Rem         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Rem         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Rem     Case Else
        Rem End Select
        
        If XCostoOrden = 0 Then
            XCostoOrden = WCosto
        End If
        
        Costo = Costo + (Cantidad * XCostoOrden * Val(XVector))
        
    Next Da
    
End Sub


