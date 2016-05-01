VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImpreCargaI 
   AutoRedraw      =   -1  'True
   Caption         =   "Impresion de Registro de Produccion"
   ClientHeight    =   6015
   ClientLeft      =   2010
   ClientTop       =   735
   ClientWidth     =   7950
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   7950
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Impre9 
         Caption         =   "Almacenero"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   5040
         Width           =   3255
      End
      Begin VB.TextBox HastaEtapa 
         Alignment       =   1  'Right Justify
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
         Left            =   4080
         MaxLength       =   4
         TabIndex        =   16
         Text            =   " "
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox DesdeEtapa 
         Alignment       =   1  'Right Justify
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
         Left            =   3000
         MaxLength       =   4
         TabIndex        =   15
         Text            =   " "
         Top             =   2880
         Width           =   855
      End
      Begin VB.CheckBox Impre8 
         Caption         =   "Lavado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   14
         Top             =   4680
         Width           =   3255
      End
      Begin VB.CheckBox Impre7 
         Caption         =   "Calidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   4320
         Width           =   3255
      End
      Begin VB.CheckBox Impre6 
         Caption         =   "Humedad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   12
         Top             =   3960
         Width           =   3255
      End
      Begin VB.CheckBox Impre5 
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   11
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CheckBox Impre4 
         Caption         =   "Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   3240
         Width           =   3255
      End
      Begin VB.CheckBox Impre3 
         Caption         =   "Procedimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   2880
         Width           =   3255
      End
      Begin VB.CheckBox Impre2 
         Caption         =   "Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   2520
         Width           =   3255
      End
      Begin VB.CheckBox Impre1 
         Caption         =   "Caratula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   2160
         Width           =   3255
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   300
         Left            =   2400
         TabIndex        =   0
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3000
         TabIndex        =   6
         Top             =   1560
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1320
         TabIndex        =   5
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
         Height          =   375
         Left            =   4320
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton AceptaII 
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
         Left            =   4320
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Producto Terminado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1815
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wficter.rpt"
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
Attribute VB_Name = "PrgImpreCargaI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String

Dim XParam As String
Dim ZDesdePaso As String
Dim ZHastaPaso As String

Dim ZHumedad(100) As String
Dim ZImpreCarga(200, 6) As String
Dim ZImpreCargaI(100, 20) As String
Dim ZImpreMetodo(100) As String
Dim ZDesTerminado As String
Dim ZFabrica As String
Dim ZZCantidad As Integer

Dim ZVector(100, 10) As String

Private Sub Acepta_Click()

    Terminado.Text = UCase(Terminado.Text)
    
    Listado.WindowTitle = "Instrucciones de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CargaI.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaI.Clave, CargaI.Terminado, CargaI.Equipo, " _
                + "Equipo.Descripcion, Equipo.DescripcionII, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaI CargaI, " _
                + DSQ + ".dbo.Equipo Equipo, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "CargaI.Equipo = Equipo.Codigo AND " _
                + "CargaI.Terminado = Terminado.Codigo AND " _
                + "CargaI.Terminado >= '" + Terminado.Text + "' AND " _
                + "CargaI.Terminado <= '" + Terminado.Text + "'"

    Listado.Connect = Connect()
    Listado.ReportFileName = "ZImpreCargaI.rpt"
    Listado.Action = 1
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Terminado.Text = UCase(Terminado.Text)
    
    Listado.WindowTitle = "Instrucciones de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CargaII.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaII.Clave, CargaII.Terminado, CargaII.MaterialAuxiliar, " _
                    + "Terminado.Descripcion, " _
                    + "MaterialAuxiliar.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaII CargaII, " _
                    + DSQ + ".dbo.Terminado Terminado, " _
                    + DSQ + ".dbo.MaterialAuxiliar MaterialAuxiliar " _
                    + "Where " _
                    + "CargaII.Terminado = Terminado.Codigo AND " _
                    + "CargaII.MaterialAuxiliar = MaterialAuxiliar.Codigo AND " _
                    + "CargaII.Terminado >= '" + Terminado.Text + "' AND " _
                    + "CargaII.Terminado <= '" + Terminado.Text + "'"

    Listado.Connect = Connect()
    Listado.ReportFileName = "ZImpreCargaII.rpt"
    Listado.Action = 1
    
    
    
    
    
    
    
    
    
    
    
    
    
    Terminado.Text = UCase(Terminado.Text)
    
    Listado.WindowTitle = "Instrucciones de Produccion"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Articulo, CargaIII.PTerminado, CargaIII.Letra, CargaIII.Descripcion, CargaIII.Cantidad, CargaIII.Cantidadii, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaIII CargaIII, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "CargaIII.Terminado = Terminado.Codigo AND " _
                + "CargaIII.Terminado >= '" + Terminado.Text + "' AND " _
                + "CargaIII.Terminado <= '" + Terminado.Text + "'"

    Listado.Connect = Connect()
    Listado.ReportFileName = "ZImpreCargaIII.rpt"
    Listado.Action = 1
    
End Sub

Private Sub AceptaII_Click()


    Terminado.Text = UCase(Terminado.Text)

    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        ZDesTerminado = rstTerminado!Descripcion
        ZFabrica = Str$(rstTerminado!fabrica)
        rstTerminado.Close
    End If

    Sql1 = "DELETE Hoja"
    Sql2 = " Where Hoja = 0"
    spHoja = Sql1 + Sql2
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = "DELETE ImpreCarga"
    spImpreCarga = ZSql
    Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = "DELETE ImpreCargaI"
    spImpreCargaI = ZSql
    Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    Erase ZVector
    Renglon = 0
    
    spComposicion = "ConsultaComposicionProducto " + "'" + Terminado.Text + "'"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstComposicion.RecordCount > 0 Then
        With rstComposicion
            .MoveFirst
            Do
                If .EOF = False Then
    
                    ZZEntraCompo = "S"
                    
                    If rstComposicion!Tipo = "M" Then
                        If Left$(UCase(rstComposicion!Articulo1), 2) = "YA" Then
                            ZZEntraCompo = "N"
                        End If
                    End If
                    
                    If ZZEntraCompo = "S" Then
    
                        Renglon = Renglon + 1
                        
                        ZVector(Renglon, 1) = rstComposicion!Tipo
                        ZVector(Renglon, 2) = rstComposicion!Articulo2
                        If rstComposicion!Articulo1 = "  -   -  " Then
                            ZVector(Renglon, 3) = "  -   -   "
                                Else
                            ZVector(Renglon, 3) = rstComposicion!Articulo1
                        End If
                        
                        ZZZZCantidad = rstComposicion!Cantidad * Val(ZFabrica)
                        ZVector(Renglon, 5) = Str$(ZZZZCantidad)
                
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstComposicion.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WTipo = ZVector(Ciclo, 1)
        WTerminado = ZVector(Ciclo, 2)
        WArticulo = ZVector(Ciclo, 3)
        WCantidad = ZVector(Ciclo, 5)
        
        Auxi1 = Str$(Ciclo)
        Call Ceros(Auxi1, 2)
    
        WClave = "000000" + Auxi1
        WHoja = "0"
        WRenglon = Auxi1
        WFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        WProducto = Terminado.Text
        WTeorico = ZFabrica
        WReal = "0"
        WFechaing = ""
        WFechaingord = ""
        Rem WTipo = "T"
        Rem WArticulo = ""
        Rem WTerminado = Terminado.Text
        Rem WCantidad = ""
        WLote = ""
        WDate = ""
        WImporte = ""
        WMarca = ""
        WSaldo = "0"
        WLote1 = ""
        WLote2 = ""
        WLote3 = ""
        WCanti1 = ""
        WCanti2 = ""
        WCanti3 = ""
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
                + WLote3 + "','" + WLote3 + "','" _
                + WCosto1 + "','" _
                + WCosto2 + "','" _
                + WCosto3 + "'"
                                           
        spHoja = "AltaHoja " + XParam
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        
        WImpreArticulo = ""
        Select Case WTipo
            Case "T"
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WImpreArticulo = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
            Case "M"
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WImpreArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
            Case Else
        End Select
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " ImpreArticulo = " + "'" + WImpreArticulo + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
    Next Ciclo
    
    
    Erase ZImpreCarga
    ZRenglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII], Equipo.Poe as [WPoe], Equipo.Identificacion as [WIdentificacion], Equipo.PoeLimpieza as [WPoeLimpieza]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "1"
                    ZImpreCarga(ZRenglon, 2) = rstCargaI!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = rstCargaI!WDescripcionII
                    ZImpreCarga(ZRenglon, 4) = rstCargaI!WPoe
                    ZImpreCarga(ZRenglon, 5) = rstCargaI!WIdentificacion
                    ZImpreCarga(ZRenglon, 6) = rstCargaI!WPoeLimpieza
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    ZSql = ""
    ZSql = ZSql + "Select *, MaterialAuxiliar.Descripcion as [WDescripcion]"
    ZSql = ZSql + " FROM CargaII, MaterialAuxiliar"
    ZSql = ZSql + " Where CargaII.MaterialAuxiliar = MaterialAuxiliar.Codigo"
    ZSql = ZSql + " and CargaII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by CargaII.Clave"
    
    rsCargaII = ZSql
    Set rstCargaII = db.OpenRecordset(rsCargaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaII.RecordCount > 0 Then
        With rstCargaII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZRenglon = ZRenglon + 1
                    
                    ZImpreCarga(ZRenglon, 1) = "2"
                    ZImpreCarga(ZRenglon, 2) = rstCargaII!WDescripcion
                    ZImpreCarga(ZRenglon, 3) = ""
                    ZImpreCarga(ZRenglon, 4) = ""
                    ZImpreCarga(ZRenglon, 5) = ""
                    ZImpreCarga(ZRenglon, 6) = ""
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaII.Close
    End If
    
    
    
    ZLugarHumedad = 0
    Erase ZHumedad
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " and CargaIII.Humedad = 1"
    ZSql = ZSql + " and CargaIII.Renglon = 1"
    ZSql = ZSql + " Order by CargaIII.Clave"
    
    rsCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        With rstCargaIII
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugarHumedad = ZLugarHumedad + 1
                    ZHumedad(ZLugarHumedad) = !Equipo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaIII.Close
    End If
    
    
    
    
    
    ZZVersion = 0
    ZZFechaVersion = ""
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaIII"
    ZSql = ZSql + " Where CargaIII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by CargaIII.Clave"
    rsCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(rsCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaIII.RecordCount > 0 Then
        ZZVersion = IIf(IsNull(rstCargaIII!Version), "", rstCargaIII!Version)
        ZZFechaVersion = IIf(IsNull(rstCargaIII!FechaVersion), "  /  /    ", rstCargaIII!FechaVersion)
        rstCargaIII.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Hoja SET "
    ZSql = ZSql + " ImpreVersion = " + "'" + Str$(ZZVersion) + "',"
    ZSql = ZSql + " ImprefechaVersion = " + "'" + ZZFechaVersion + "'"
    ZSql = ZSql + " Where Hoja = " + "'" + "0" + "'"
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
    
    For ZCiclo = 1 To 100
        
        ZTipo = ZImpreCarga(ZCiclo, 1)
        ZDescripcion = ZImpreCarga(ZCiclo, 2)
        ZDescripcionII = ZImpreCarga(ZCiclo, 3)
        ZPoe = Trim(ZImpreCarga(ZCiclo, 4))
        ZIdentificacion = Trim(ZImpreCarga(ZCiclo, 5))
        ZPoeLimpieza = Trim(ZImpreCarga(ZCiclo, 6))
        If ZIdentificacion <> "" Then
            ZDescripcion = ZIdentificacion + " - " + ZDescripcion
        End If
        
        
        If ZTipo <> "" Then
                                
            ZSql = ""
            ZSql = ZSql & "INSERT INTO ImpreCarga ("
            ZSql = ZSql & "Partida ,"
            ZSql = ZSql & "Descripcion ,"
            ZSql = ZSql & "Terminado ,"
            ZSql = ZSql & "Cantidad ,"
            ZSql = ZSql & "Tipo ,"
            ZSql = ZSql & "DescripcionI ,"
            ZSql = ZSql & "DescripcionII )"
            ZSql = ZSql & "Values ("
            ZSql = ZSql & "'" + "0" + "',"
            ZSql = ZSql & "'" + ZDesTerminado + "',"
            ZSql = ZSql & "'" + Terminado.Text + "',"
            ZSql = ZSql & "'" + ZFabrica + "',"
            ZSql = ZSql & "'" + ZTipo + "',"
            ZSql = ZSql & "'" + ZDescripcion + "',"
            ZSql = ZSql & "'" + ZDescripcionII + "')"
        
            spImpreCarga = ZSql
            Set rstImpreCarga = db.OpenRecordset(spImpreCarga, dbOpenSnapshot, dbSQLPassThrough)
        
        End If
                            
    Next ZCiclo
    
    
    
    
    
    
    
    
    
    
    Erase ZImpreMetodo
    ZLugarMetodo = 0
    Erase ZImpreCargaI
    ZRenglon = 0
    ZLugar = 1
    
    ZSql = ""
    ZSql = ZSql + "Select *, Equipo.Descripcion as [WDescripcion], Equipo.DescripcionII as [WDescripcionII], Equipo.Poe as [WPoe], Equipo.Identificacion as [WIdentificacion], Equipo.PoeLimpieza as [WPoeLimpieza]"
    ZSql = ZSql + " FROM CargaI, Equipo"
    ZSql = ZSql + " Where CargaI.Equipo = Equipo.Codigo"
    ZSql = ZSql + " and CargaI.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by CargaI.Clave"
    
    rsCargaI = ZSql
    Set rstCargaI = db.OpenRecordset(rsCargaI, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaI.RecordCount > 0 Then
        With rstCargaI
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZEquipo = rstCargaI!Equipo
                    ZZDescripcionI = rstCargaI!WDescripcion
                    ZZDescripcionII = rstCargaI!WDescripcionII
                    ZZZMetodo = Trim(rstCargaI!WPoeLimpieza) + " - " + IIf(IsNull(rstCargaI!Metodo), "", rstCargaI!Metodo)
                    ZZCantidad = IIf(IsNull(rstCargaI!Cantidad), "0", rstCargaI!Cantidad)
                    ZZPoe = rstCargaI!WPoe
                    ZZIdentificacion = rstCargaI!WIdentificacion
                    ZZPoeLimpieza = rstCargaI!WPoeLimpieza
                    If Trim(ZZIdentificacion) <> "" Then
                        ZZDescripcionI = Trim(ZZIdentificacion) + " - " + ZZDescripcionI
                    End If
                    
                    Rem If ZZCantidad <> 0 Then
                    Rem
                    Rem     ZEntraMetodo = "S"
                    Rem
                    Rem     For ZCicloMetodo = 1 To ZLugarMetodo
                    Rem         If ZImpreMetodo(ZCicloMetodo) = Trim(ZZZMetodo) Then
                    Rem             ZEntraMetodo = "N"
                    Rem             Exit For
                    Rem         End If
                    Rem     Next ZCicloMetodo
                    Rem
                    Rem     If ZEntraMetodo = "S" Then
                    Rem         ZLugarMetodo = ZLugarMetodo + 1
                    Rem         ZImpreMetodo(ZLugarMetodo) = Trim(ZZZMetodo)
                    Rem     End If
                    Rem
                    Rem End If
                    
                    For Ciclo = 1 To ZZCantidad
                        Select Case ZLugar
                            Case 1
                                ZRenglon = ZRenglon + 1
                                ZImpreCargaI(ZRenglon, 1) = Terminado.Text
                                ZImpreCargaI(ZRenglon, 2) = ZDesTerminado
                                ZImpreCargaI(ZRenglon, 3) = "0"
                                ZImpreCargaI(ZRenglon, 4) = "0"
                                
                                ZImpreCargaI(ZRenglon, 5) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 6) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 7) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 8) = ZZZMetodo
                                
                                ZLugar = 2
                            
                            Case 2
                                ZImpreCargaI(ZRenglon, 9) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 10) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 11) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 12) = ZZZMetodo
                            
                                ZLugar = 3
                                
                            Case 3
                                ZImpreCargaI(ZRenglon, 13) = ZZEquipo
                                ZImpreCargaI(ZRenglon, 14) = ZZDescripcionI
                                ZImpreCargaI(ZRenglon, 15) = ZZDescripcionII
                                ZImpreCargaI(ZRenglon, 16) = ZZZMetodo
                                
                                ZLugar = 1
                            Case Else
                        End Select
                    Next Ciclo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCargaI.Close
    End If
    
    
    
    For ZCiclo = 1 To ZRenglon
    
        ZZCodigo = Str$(ZCiclo)
        
        ZZTerminado = ZImpreCargaI(ZCiclo, 1)
        ZZDescripcion = ZImpreCargaI(ZCiclo, 2)
        ZZPartida = ZImpreCargaI(ZCiclo, 3)
        ZZCantidad = ZImpreCargaI(ZCiclo, 4)
                                
        ZZEquipoI = ZImpreCargaI(ZCiclo, 5)
        ZZDesEquipoI = ZImpreCargaI(ZCiclo, 6)
        ZZDesEquipoOtroI = ZImpreCargaI(ZCiclo, 7)
        ZZMetodoI = ZImpreCargaI(ZCiclo, 8)
                                
        ZZEquipoII = ZImpreCargaI(ZCiclo, 9)
        ZZDesEquipoII = ZImpreCargaI(ZCiclo, 10)
        ZZDesEquipoOtroII = ZImpreCargaI(ZCiclo, 11)
        ZZMetodoII = ZImpreCargaI(ZCiclo, 12)
                                
        ZZEquipoIII = ZImpreCargaI(ZCiclo, 13)
        ZZDesEquipoIII = ZImpreCargaI(ZCiclo, 14)
        ZZDesEquipoOtroIII = ZImpreCargaI(ZCiclo, 15)
        ZZMetodoIII = ZImpreCargaI(ZCiclo, 16)
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO ImpreCargaI ("
        ZSql = ZSql & "Codigo ,"
        ZSql = ZSql & "Terminado ,"
        ZSql = ZSql & "Descripcion ,"
        ZSql = ZSql & "Partida ,"
        ZSql = ZSql & "Cantidad ,"
        ZSql = ZSql & "MetodoI ,"
        ZSql = ZSql & "EquipoI ,"
        ZSql = ZSql & "DesEquipoI ,"
        ZSql = ZSql & "DesEquipoOtroI ,"
        ZSql = ZSql & "MetodoII ,"
        ZSql = ZSql & "EquipoII ,"
        ZSql = ZSql & "DesEquipoII ,"
        ZSql = ZSql & "DesEquipoOtroII ,"
        ZSql = ZSql & "MetodoIII ,"
        ZSql = ZSql & "EquipoIII ,"
        ZSql = ZSql & "DesEquipoIII ,"
        ZSql = ZSql & "DesEquipoOtroIII )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + ZZCodigo + "',"
        ZSql = ZSql & "'" + ZZTerminado + "',"
        ZSql = ZSql & "'" + ZZDescripcion + "',"
        ZSql = ZSql & "'" + ZZPartida + "',"
        ZSql = ZSql & "'" + Str$(ZZCantidad) + "',"
        ZSql = ZSql & "'" + ZZMetodoI + "',"
        ZSql = ZSql & "'" + ZZEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoI + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroI + "',"
        ZSql = ZSql & "'" + ZZMetodoII + "',"
        ZSql = ZSql & "'" + ZZEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroII + "',"
        ZSql = ZSql & "'" + ZZMetodoIII + "',"
        ZSql = ZSql & "'" + ZZEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoIII + "',"
        ZSql = ZSql & "'" + ZZDesEquipoOtroIII + "')"
        
        spImpreCargaI = ZSql
        Set rstImpreCargaI = db.OpenRecordset(spImpreCargaI, dbOpenSnapshot, dbSQLPassThrough)
                            
    Next ZCiclo
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaIII SET "
    ZSql = ZSql + " Partida = " + "'" + "0" + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + ZFabrica + "'"
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaIII = ZSql
    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CargaV SET "
    ZSql = ZSql + " Partida = " + "'" + "0" + "',"
    ZSql = ZSql + " CantidadPartida = " + "'" + ZFabrica + "',"
    ZSql = ZSql + " ImprePaso = Paso "
    ZSql = ZSql + " Where Terminado = " + "'" + Terminado.Text + "'"
    spCargaV = ZSql
    Set rstCargaV = db.OpenRecordset(spCargaV, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
    
  Rem nan


    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.Connect = Connect()
    
    If Impre1.Value = 1 Then
    
        XEmpresa = Wempresa
        Wempresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        ZZLoteAutoriza = ""
        ZZFabrica = ""
        ZZfabricaII = ""
            
        Sql1 = "Select *"
        Sql2 = " FROM Terminado"
        Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            ZZLoteAutorizado = rstTerminado!loteautorizado
            ZZfabricaII = IIf(IsNull(rstTerminado!fabricaII), "0", rstTerminado!fabricaII)
            ZZfabricaIII = IIf(IsNull(rstTerminado!fabricaIII), "0", rstTerminado!fabricaIII)
            rstTerminado.Close
        End If
        
        Call Conecta_Empresa
        
        If ZZfabricaII <> 0 And ZZfabricaIII <> 0 Then
            WLoteAutorizado = Trim(Str$(ZZfabricaII)) + " a " + Trim(Str$(ZZfabricaIII))
                Else
            WLoteAutorizado = ""
        End If
        
        Sql1 = "UPDATE Terminado SET "
        Sql2 = " LoteAutorizado = " + "'" + WLoteAutorizado + "'"
        Sql3 = " Where Codigo = " + "'" + Terminado.Text + "'"
        spTerminado = Sql1 + Sql2 + Sql3
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, Hoja.ImpreVersion, Hoja.ImpreFechaVersion, " _
                    + "Terminado.Descripcion, Terminado.Version, Terminado.FechaVersion, Terminado.LoteAutorizado " _
                    + "From " _
                    + DSQ + ".dbo.Hoja Hoja, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "Hoja.Producto = Terminado.Codigo AND " _
                    + "Hoja.Hoja >= " + "0" + " AND " _
                    + "Hoja.Hoja <= " + "0"
    
        Listado.ReportFileName = "WImpreCaratula.rpt"
        Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
        Listado.SelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
      Rem  Listado.Destination = 1
        Listado.Action = 1
    
    End If
    
    
    
    If Impre2.Value = 1 Then
    
    Listado.SQLQuery = "SELECT ImpreCarga.Partida, ImpreCarga.Terminado, ImpreCarga.Descripcion, ImpreCarga.Cantidad, ImpreCarga.Tipo, ImpreCarga.DescripcionI, ImpreCarga.DescripcionII " _
                + "From " _
                + DSQ + ".dbo.ImpreCarga ImpreCarga " _
                + "Where " _
                + "ImpreCarga.Partida >= " + "0" + " AND " _
                + "ImpreCarga.Partida <= " + "0"
    
    Listado.ReportFileName = "ImpreEquipos.rpt"
    Listado.GroupSelectionFormula = "{ImpreCarga.Partida} in " + "0" + " to " + "0"
    Listado.SelectionFormula = "{ImpreCarga.Partida} in " + "0" + " to " + "0"
   Rem Listado.Destination = 1
    Listado.Action = 1
    
    End If
    
    

    If Impre3.Value = 1 Then
    
    If Val(DesdeEtapa.Text) = 0 And Val(HastaEtapa.Text) = 0 Then
        ZDesdePaso = "0"
        ZHastaPaso = "999"
            Else
        ZDesdePaso = DesdeEtapa.Text
        ZHastaPaso = HastaEtapa.Text
    End If
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Renglon, CargaIII.Articulo, CargaIII.PTerminado, CargaIII.Letra, CargaIII.Descripcion, CargaIII.Cantidad, CargaIII.Partida, CargaIII.CantidadPartida , " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.CargaIII CargaIII, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "CargaIII.Terminado = Terminado.Codigo AND " _
                    + "CargaIII.Terminado >= '" + Terminado.Text + "' AND " _
                    + "CargaIII.Terminado <= '" + Terminado.Text + "' AND " _
                    + "CargaIII.Paso >= " + ZDesdePaso + " AND " _
                    + "CargaIII.Paso <= " + ZHastaPaso

    Listado.ReportFileName = "WImpreProcedimiento.rpt"
    
    Uno = "{CargaIII.Paso} in " + ZDesdePaso + " to " + ZHastaPaso
    Dos = " and {CargaIII.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    
    Listado.GroupSelectionFormula = Uno + Dos
    Listado.SelectionFormula = Uno + Dos
    
  Rem  Listado.Destination = 1
    Listado.Action = 1
    
    End If
    
    
    
        
    

    
    If Impre4.Value = 1 Then
    
    Listado.SQLQuery = "SELECT CargaIII.Clave, CargaIII.Terminado, CargaIII.Paso, CargaIII.Partida, CargaIII.CantidadPartida, CargaIII.Peso, CargaIII.ImprePeso, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.CargaIII CargaIII, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "CargaIII.Terminado = Terminado.Codigo AND " _
                + "CargaIII.Terminado >= '" + Terminado.Text + "' AND " _
                + "CargaIII.Terminado <= '" + Terminado.Text + "' AND " _
                + "CargaIII.Peso = 1 AND " _
                + "CargaIII.ImprePeso = 'S'"

    Listado.ReportFileName = "ImprePeso.rpt"
    Listado.GroupSelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
    Listado.SelectionFormula = "{CargaIII.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
  Rem  Listado.Destination = 1
    Listado.Action = 1
    
    End If


    If Impre5.Value = 1 Then
    
    Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
            + "Terminado.Descripcion " _
            + "From " _
            + DSQ + ".dbo.Hoja Hoja, " _
            + DSQ + ".dbo.Terminado Terminado " _
            + "Where " _
            + "Hoja.Producto = Terminado.Codigo AND " _
            + "Hoja.Hoja >= " + "0" + " AND " _
            + "Hoja.Hoja <= " + "0"
            
    Listado.ReportFileName = "ImpreObservaciones.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
    Listado.SelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
  Rem  Listado.Destination = 1
    Listado.Action = 1
    Listado.Action = 1
    
    End If
    
    
    
    
    
    
    
    
    If Impre6.Value = 1 Then
    
    For CicloHumedad = 1 To ZLugarHumedad
    
        WIdentificacion = ""
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + ZHumedad(CicloHumedad) + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            WIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            rstEquipo.Close
        End If
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Hoja SET "
        ZSql = ZSql + " Identificacion = " + "'" + WIdentificacion + "'"
        ZSql = ZSql + " Where Hoja = " + "'" + "0" + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        Listado.SQLQuery = "SELECT Hoja.Hoja, Hoja.Producto, Hoja.Teorico, " _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + "0" + " AND " _
                + "Hoja.Hoja <= " + "0"

        Listado.ReportFileName = "ImpreHumedad.rpt"
        Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
        Listado.SelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
     Rem   Listado.Destination = 1
        Listado.Action = 1
        
    Next CicloHumedad
    
    End If
    
    
    
    
    
        If Impre7.Value = 1 Then
        
        Listado.SQLQuery = "SELECT CargaV.Clave, CargaV.Terminado, CargaV.Paso, CargaV.Valor, CargaV.Ensayo, CargaV.DesEnsayo, CargaV.Partida, CargaV.CantidadPartida, CargaV.Corte, CargaV.ImprePaso, " _
                        + "Terminado.Descripcion " _
                        + "From " _
                        + DSQ + ".dbo.CargaV CargaV, " _
                        + DSQ + ".dbo.Terminado Terminado " _
                        + "Where " _
                        + "CargaV.Terminado = Terminado.Codigo AND " _
                        + "CargaV.Terminado >= '" + Terminado.Text + "' AND " _
                        + "CargaV.Terminado <= '" + Terminado.Text + "'"
    
        Listado.ReportFileName = "ImpreCalidad.rpt"
        Listado.GroupSelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
        Listado.SelectionFormula = "{CargaV.Terminado} in " + Chr$(34) + Terminado.Text + Chr$(34) + " to " + Chr$(34) + Terminado.Text + Chr$(34)
       Rem Listado.Destination = 1
        Listado.Action = 1
        
        End If
    
    Rem If Impre8.Value = 1 Then
    Rem
    Rem Listado.SQLQuery = "SELECT ImpreCargaI.Codigo, ImpreCargaI.Terminado, ImpreCargaI.Descripcion, ImpreCargaI.Partida, ImpreCargaI.Cantidad, ImpreCargaI.MetodoI, ImpreCargaI.EquipoI, ImpreCargaI.DesEquipoI, ImpreCargaI.DesEquipoOtroI, ImpreCargaI.MetodoII, ImpreCargaI.EquipoII, ImpreCargaI.DesEquipoII, ImpreCargaI.DesEquipoOtroII, ImpreCargaI.MetodoIII, ImpreCargaI.EquipoIII, ImpreCargaI.DesEquipoIII, ImpreCargaI.DesEquipoOtroIII " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ImpreCargaI ImpreCargaI " _
    rem             + "Where " _
    rem             + "ImpreCargaI.Codigo >= 0 AND " _
    rem             + "ImpreCargaI.Codigo <= 999999"
    Rem
    Rem Listado.ReportFileName = "ImpreIdentificacion.rpt"
    Rem Listado.GroupSelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Listado.SelectionFormula = "{ImpreCargaI.Codigo} in 0 to 999999"
    Rem Rem Listado.Destination = 1
    Rem Listado.Action = 1
    Rem
    Rem End If
    
    
    
    
    
    
    If Impre9.Value = 1 Then
    
    Listado.SQLQuery = "SELECT Hoja.Clave, Hoja.Hoja, Hoja.Producto, Hoja.Cantidad, Hoja.Tipo, Hoja.Articulo, Hoja.Terminado, Hoja.Teorico, Hoja.ImpreArticulo," _
                + "Terminado.Descripcion " _
                + "From " _
                + DSQ + ".dbo.Hoja Hoja, " _
                + DSQ + ".dbo.Terminado Terminado " _
                + "Where " _
                + "Hoja.Producto = Terminado.Codigo AND " _
                + "Hoja.Hoja >= " + "0" + " AND " _
                + "Hoja.Hoja <= " + "0"

    Listado.ReportFileName = "ImpreHojaFarmaAlmacen.rpt"
    Listado.GroupSelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
    Listado.SelectionFormula = "{Hoja.Hoja} in " + "0" + " to " + "0"
    Rem Listado.Destination = 1
    Listado.Action = 1
    
    End If
    
    

    Sql1 = "DELETE Hoja"
    Sql2 = " Where Hoja = 0"
    spHoja = Sql1 + Sql2
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    


End Sub

Private Sub CANCELA_Click()
    Terminado.SetFocus
    PrgImpreCargaI.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Impre1.Value = 0
    Impre2.Value = 0
    Impre3.Value = 0
    Impre4.Value = 0
    Impre5.Value = 0
    Impre6.Value = 0
    Impre7.Value = 0
    Impre8.Value = 0

    Panta.Value = False
    Impresora.Value = True
End Sub


Private Sub Conecta_Empresa()

    Select Case Val(XEmpresa)
        Case 1
            Wempresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            Wempresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            Wempresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            Wempresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            Wempresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            Wempresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            Wempresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            Wempresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            Wempresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select

End Sub

