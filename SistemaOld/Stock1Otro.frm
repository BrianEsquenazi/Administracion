VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgStock1Otro 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valorizacion de Materia Prima a Fecha"
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
      Height          =   3135
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   4815
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
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
         Left            =   3600
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
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
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   1095
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
         Left            =   960
         TabIndex        =   5
         Top             =   2400
         Width           =   1215
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
         Left            =   2400
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
         Caption         =   "Tipo"
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
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1215
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
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
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
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PrgStock1Otro"
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
Private Vector(10000, 3) As String
Dim Empe(12, 10) As String
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

Dim Costo1 As Double
Dim WCosto1 As Double
Dim ZCosto1 As Double
Dim ZZParidad As Double
Dim ZZCostoArti As Double

Dim ZZOrdenI As Double
Dim ZZPtaOrdenI As Integer
Dim ZZFechaOrdenI As String
Dim ZZFechaOrdI As String

Dim ZZOrdenII As Double
Dim ZZPtaOrdenII As Integer
Dim ZZFechaOrdenII As String
Dim ZZFechaOrdII As String

Dim ZZOrdenIII As Double
Dim ZZPtaOrdenIII As Integer
Dim ZZFechaOrdenIII As String
Dim ZZFechaOrdIII As String

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

    Erase Vector
    Renglon = 0

    spArticulo = "ListaArticuloStock"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstArticulo

            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstArticulo!Codigo >= Desde.Text And rstArticulo!Codigo <= Hasta.Text Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstArticulo!Codigo
                    WStock = Str$(rstArticulo!Entradas - rstArticulo!Salidas)
                    Vector(Renglon, 2) = WStock
                    Vector(Renglon, 3) = Str$(rstArticulo!Costo2)
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
    
    spArticulo = "ModificaArticuloStock0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WArticulo = Vector(Da, 1)
        XCodigo = Vector(Da, 1)
        XStock = Val(Vector(Da, 2))
        WCosto = Val(Vector(Da, 3))
        XDate = Date$
        
        Rem If WArticulo = "CO-227-100" Then Stop
        
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
        
        If Impo4 > 0 And Tipo.ListIndex = 1 Then
            Call Calcula_Costo
            If ZZCostoArti <> 0 Then
                WCosto = ZZCostoArti
            End If
        End If
        
        
        Rem If Impo4 > 0 Then
            WStock = Str$(Impo4)
            XCostoOrden = WCosto
            XParam = "'" + XCodigo + "','" _
                    + WStock + "','" _
                    + Str$(XCostoOrden) + "'"
            spArticulo = "ModificaArticuloStock " + XParam
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        Rem End If
        
    Next Da
    
    Listado.WindowTitle = "Listado de Valorizacion de Materia Prima a Fecha"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaEnv.Envase} in " + DesdeEnv.Text + " to " + HastaEnv.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Articulo.Codigo, Articulo.Descripcion, Articulo.Costo, Articulo.Stock " _
                        + "From " _
                        + DSQ + ".dbo.Articulo Articulo " _
                        + "Where " _
                        + "Articulo.Codigo >= '  -   -   ' AND " _
                        + "Articulo.Codigo <= 'ZZ-999-999' AND Articulo.Stock <> 0."
    
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
                        WLiberada = IIf(IsNull(rstLaudo!Liberadaant), 0, rstLaudo!Liberadaant)
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
        rstMovlab.Close
    End If
    
    
    Rem PROCESA LAS VENTAS
    
    If Left$(WArticulo, 2) = "DY" Or Left$(WArticulo, 2) = "DS" Or Left$(WArticulo, 2) = "DQ" Then
    
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
        rstEstadistica.Close
    End If
    
    End If

End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Fecha.SetFocus
    PrgStock1Otro.Hide
    Unload Me
    Menu.Show
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

    Tipo.Clear
    
    Tipo.AddItem "Costo Std"
    Tipo.AddItem "Costo Ult.Compra"
    
    Tipo.ListIndex = 0
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgStock1Otro.Caption = "Listado de Valorizacion de Materia Prima a Fecha :  " + !Nombre
        End If
    End With
    
    Fecha.Text = "  /  /    "
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Sub Calcula_Costo()

    XEmpresa = WEmpresa
    ZZCostoArti = 0

    spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        
        Costo1 = rstArticulo!Costo1
        WCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
        ZCosto1 = IIf(IsNull(rstArticulo!ZCosto1), "0", rstArticulo!ZCosto1)
        
        ZZOrdenI = IIf(IsNull(rstArticulo!OrdenI), "0", rstArticulo!OrdenI)
        ZZOrdenII = IIf(IsNull(rstArticulo!OrdenII), "0", rstArticulo!OrdenII)
        ZZOrdenIII = IIf(IsNull(rstArticulo!OrdenIII), "0", rstArticulo!OrdenIII)
        ZZPtaOrdenI = IIf(IsNull(rstArticulo!PtaOrdenI), "0", rstArticulo!PtaOrdenI)
        ZZPtaOrdenII = IIf(IsNull(rstArticulo!PtaOrdenII), "0", rstArticulo!PtaOrdenII)
        ZZPtaOrdenIII = IIf(IsNull(rstArticulo!PtaOrdenIII), "0", rstArticulo!PtaOrdenIII)
        
        rstArticulo.Close
        
        ZZFechaOrdenI = ""
        ZZFechaOrdenII = ""
        ZZFechaOrdenIII = ""
        
        If ZZPtaOrdenI <> 0 And ZZOrdenI <> 0 Then
        
            ZZImpre = ""
            
            Select Case ZZPtaOrdenI
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    WEmpresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + Str(ZZOrdenI) + "'"
            ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                ZZFechaOrdenI = rstOrden!Fecha
                rstOrden.Close
            End If
            
            Call Conecta_Empresa
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        If ZZPtaOrdenII <> 0 And ZZOrdenII <> 0 Then
        
            ZZImpre = ""
            
            Select Case ZZPtaOrdenII
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    WEmpresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenII) + "'"
            ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                ZZFechaOrdenII = rstOrden!Fecha
                rstOrden.Close
            End If
            
            Call Conecta_Empresa
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        If ZZPtaOrdenIII <> 0 And ZZOrdenIII <> 0 Then
        
            ZZImpre = ""
            
            Select Case ZZPtaOrdenIII
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SI"
                Case 2
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PI"
                Case 3
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SII"
                Case 4
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PII"
                Case 5
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIII"
                Case 6
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SIV"
                Case 7
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SV"
                Case 8
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PIII"
                Case 9
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "PV"
                Case 10
                    WEmpresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVI"
                Case 11
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    ZZImpre = "SVII"
                Case Else
            End Select
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Orden"
            ZSql = ZSql + " Where Orden = " + "'" + Str$(ZZOrdenIII) + "'"
            ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
            spOrden = ZSql
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                ZZFechaOrdenIII = rstOrden!Fecha
                rstOrden.Close
            End If
            
            Call Conecta_Empresa
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
                
        If ZZFechaOrdenI <> "" Then
            ZZFechaOrdI = Right$(ZZFechaOrdenI, 4) + Mid$(ZZFechaOrdenI, 4, 2) + Left$(ZZFechaOrdenI, 2)
                Else
            ZZFechaOrdI = ""
        End If
        If ZZFechaOrdenII <> "" Then
            ZZFechaOrdII = Right$(ZZFechaOrdenII, 4) + Mid$(ZZFechaOrdenII, 4, 2) + Left$(ZZFechaOrdenII, 2)
                Else
            ZZFechaOrdII = ""
        End If
        If ZZFechaOrdenIII <> "" Then
            ZZFechaOrdIII = Right$(ZZFechaOrdenIII, 4) + Mid$(ZZFechaOrdenIII, 4, 2) + Left$(ZZFechaOrdenIII, 2)
                Else
            ZZFechaOrdIII = ""
        End If
        
        If ZZFechaOrdI <> "" And ZZFechaOrdI > ZZFechaOrdII And ZZFechaOrdI > ZZFechaOrdIII Then
            ZZCostoArti = Costo1
        End If
        
        If ZZFechaOrdII <> "" And ZZFechaOrdII > ZZFechaOrdI And ZZFechaOrdII > ZZFechaOrdIII Then
            ZZCostoArti = WCosto1
            
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            spCambios = "ConsultaCambio  " + "'" + ZZFechaOrdenII + "'"
            Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
            If rstCambios.RecordCount > 0 Then
                ZZParidad = rstCambios!Cambio
                rstCambios.Close
                If ZZParidad <> 0 Then
                    ZZCostoArti = WCosto1 / ZZParidad
                End If
            End If
        
        End If
        
        If ZZFechaOrdIII <> "" And ZZFechaOrdIII > ZZFechaOrdI And ZZFechaOrdIII > ZZFechaOrdII Then
            ZZCostoArti = ZCosto1
        End If
        
    End If
    
    Call Conecta_Empresa
    
End Sub

    

