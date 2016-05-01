VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListinfImporta 
   AutoRedraw      =   -1  'True
   Caption         =   "Analisis de Ordenes de Compra de Importacion"
   ClientHeight    =   6240
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   6240
   ScaleWidth      =   8145
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   5655
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
         Height          =   345
         Left            =   4200
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox HastaProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox DesdeProv 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2280
         MaxLength       =   11
         TabIndex        =   12
         Top             =   1320
         Width           =   1575
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   10
         Top             =   840
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
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   480
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
         Mask            =   "##/##/####"
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
         Height          =   375
         Left            =   3000
         TabIndex        =   9
         Top             =   2160
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
         Left            =   1440
         TabIndex        =   8
         Top             =   2160
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
         Left            =   4200
         TabIndex        =   7
         Top             =   600
         Width           =   1095
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
         Left            =   4200
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Proveedor"
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Proveedor"
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
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
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
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinfImporta.rpt"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      ItemData        =   "listinfimporta.frx":0000
      Left            =   120
      List            =   "listinfimporta.frx":0007
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   7815
   End
End
Attribute VB_Name = "PrgListinfImporta"
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
Dim WVector(1000, 4) As String
Dim WDevuelta As String
Dim WLiberada As String
Dim WPartida1 As String
Dim WPartida2 As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    Rem Calcula las diferencias de fecha entre la
    Rem Orden de compra y el informe de recepcion
    
    Rem Sql1 = "UPDATE Orden SET "
    Rem Sql2 = " TipoSolicitud = 0"
    Rem spInsumo = Sql1 + Sql2
    Rem Set rstInsumo = db.OpenRecordset(spInsumo, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase WVector
    Lugar = 0
                    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "','" _
                 + DesdeProv.Text + "','" _
                 + HastaProv.Text + "'"

    spInforme = "ListaInformeListado" + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
            
        With rstInforme
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Lugar = Lugar + 1
                
                WVector(Lugar, 1) = rstInforme!Articulo
                WVector(Lugar, 2) = rstInforme!Orden
                WVector(Lugar, 3) = rstInforme!FechaOrd
                WVector(Lugar, 4) = rstInforme!Clave
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstInforme.Close
        
    End If
    
    For Ciclo = 1 To Lugar
    
        WArticulo = WVector(Ciclo, 1)
        WOrden = WVector(Ciclo, 2)
        WFecha = WVector(Ciclo, 3)
        WClave = WVector(Ciclo, 4)
        XFecha = "  /  /    "
        XOrdFecha = "00000000"
        
        XParam = "'" + WOrden + "','" _
                + WArticulo + "'"

        spOrden = "ListaOrdenArticulo" + XParam
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
        
            WCarpeta = IIf(IsNull(rstOrden!Carpeta), "0", rstOrden!Carpeta)
            
            If WCarpeta <> 0 Then
                XOrdFecha = Right$(rstOrden!Fecha, 4) + Mid$(rstOrden!Fecha, 4, 2) + Left$(rstOrden!Fecha, 2)
                XFecha = rstOrden!fecha2
                WFechaImpo = IIf(IsNull(rstOrden!FechaImpo), "  /  /    ", rstOrden!FechaImpo)
                If WFechaImpo <> "  /  /    " Then
                    XOrdFecha = Right$(rstOrden!FechaImpo, 4) + Mid$(rstOrden!FechaImpo, 4, 2) + Left$(rstOrden!FechaImpo, 2)
                    XFecha = rstOrden!FechaImpo
                End If
                    Else
                XOrdFecha = "0000000"
                XFecha = "  /  /    "
            End If
            
            rstOrden.Close
            
        End If
        
        BAse1 = (Val(Left$(XOrdFecha, 4)) * 365) + (Val(Mid$(XOrdFecha, 5, 2)) * 30) + (Val(Right$(XOrdFecha, 2)) * 1)
        Base2 = (Val(Left$(WFecha, 4)) * 365) + (Val(Mid$(WFecha, 5, 2)) * 30) + (Val(Right$(WFecha, 2)) * 1)
        
        Dife = Base2 - BAse1
        XDife = Str$(Dife)
        
        XParam = "'" + WClave + "','" _
                + XFecha + "','" _
                + XOrdFecha + "','" _
                + XDife + "'"

        spInforme = "ModificaInformeListado" + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)

    Next Ciclo
    
    Rem Calcula las Cantidades Liberadas y Devueltas
    Rem y sus partidas respectivas
    
    Erase WVector
    Lugar = 0
 
    XParam = "'" + WDesde + "','" _
                 + WHasta + "','" _
                 + DesdeProv.Text + "','" _
                 + HastaProv.Text + "'"

    spInforme = "ListaInformeListado" + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
    
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                    Lugar = Lugar + 1
                    WVector(Lugar, 1) = rstInforme!Clave
                    WVector(Lugar, 2) = rstInforme!Informe
                    WVector(Lugar, 3) = rstInforme!Articulo
                    WVector(Lugar, 4) = rstInforme!Cantidad
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To Lugar
    
        WClave = WVector(Ciclo, 1)
        WInforme = WVector(Ciclo, 2)
        WArticulo = WVector(Ciclo, 3)
        WCantidad = Val(WVector(Ciclo, 4))
        WLiberada = ""
        WDevuelta = ""
        WPartida1 = ""
        WPartida2 = ""
        
        XParam = "'" + WInforme + "','" _
                 + WArticulo + "'"
        spLaudo = "ListaLaudoInforme " + XParam
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
                        
                            If Val(rstLaudo!Liberadaant) <> 0 Then
                                WLiberada = Str$(Val(WLiberada) + rstLaudo!Liberadaant)
                                WPartida1 = Str$(rstLaudo!Laudo)
                            End If
                            If Val(rstLaudo!devueltaant) <> 0 Then
                                WDevuelta = Str$(Val(WDevuelta) + rstLaudo!devueltaant)
                                WPartida2 = Str$(rstLaudo!Laudo)
                            End If
                            
                                Else
                                
                            If Val(rstLaudo!Liberada) <> 0 Then
                                WLiberada = Str$(Val(WLiberada) + rstLaudo!Liberada)
                                WPartida1 = Str$(rstLaudo!Laudo)
                            End If
                            If Val(rstLaudo!devuelta) <> 0 Then
                                WDevuelta = Str$(Val(WDevuelta) + rstLaudo!devuelta)
                                WPartida2 = Str$(rstLaudo!Laudo)
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
        
        XParam = "'" + WClave + "','" _
                + WLiberada + "','" _
                + WPartida1 + "','" _
                + WDevuelta + "','" _
                + WPartida2 + "'"

        spInforme = "ModificaInformeListadoII " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo

    Listado.WindowTitle = "Analisis de Ordenes de Compra de Importacion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34) + "and {Informe.ordfechaorden} <> " + Chr$(34) + "0000000" + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Informe.Informe, Informe.Fecha, Informe.Remito, Informe.Proveedor, Informe.Orden, Informe.Articulo, Informe.Cantidad, Informe.Fechaord, Informe.FechaOrden, Informe.Difefecha, " _
                    + "Articulo.Descripcion, " _
                    + "Proveedor.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Informe Informe, " _
                    + DSQ + ".dbo.Articulo Articulo, " _
                    + DSQ + ".dbo.Proveedor Proveedor " _
                    + "Where " _
                    + "Informe.Articulo = Articulo.Codigo AND Informe.Proveedor = Proveedor.Proveedor AND " _
                    + "Informe.Proveedor >= '" + DesdeProv.Text + "' AND Informe.Proveedor <= '" + HastaProv.Text + "' AND " _
                    + "Informe.Fechaord >= '" + WDesde + "' AND Informe.Fechaord <= '" + WHasta + "'" _
                    + " and Informe.OrdFechaOrden <> '00000000'"
                        
    Listado.DataFiles(3) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgListinfImporta.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String
    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spProveedor = "ListaProveedoresOrd"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
            With RstProveedor
                .MoveFirst
                Do
                    If .EOF = False Then
                        Auxi = RstProveedor!Proveedor
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = RstProveedor!Proveedor
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            RstProveedor.Close
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub


Private Sub Document1_GotFocus()

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WProveedor = WIndice.List(Indice)
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
            
            If WPasa = "S" Then
                DesdeProv.Text = WProveedor
                HastaProv.Text = WProveedor
            End If
            DesdeProv.SetFocus
            
    End Select
    
End Sub



Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
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

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeProv.SetFocus
    End If
End Sub

Private Sub DesdeProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProv.SetFocus
    End If
End Sub

Private Sub HastaProv_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgListinfImporta.Caption = "Listado de Informe de Recepcion :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spProveedor = "ListaProveedoresOrd"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                Da = Len(RstProveedor!Nombre) - WEspacios
                
                For aa = 1 To Da
                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                        Auxi = Str$(RstProveedor!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + " " + RstProveedor!Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    RstProveedor.Close
    
    End If

End Sub


