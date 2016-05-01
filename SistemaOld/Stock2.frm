VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgStock2 
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
      Caption         =   "Control Listado"
      Height          =   2535
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   5415
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
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
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Producto"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "PrgStock2"
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
Private Producto As String
Private Costo As Double
Private Auxiliar(100, 7) As String
Private WVector(10000) As String
Private WCodigo As String

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Fecha.SetFocus
    PrgStock2.Hide
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
                    WVector(Renglon) = rstTerminado!Codigo
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
        WTerminado = WVector(Da)
        XCodigo = WVector(Da)
        WCodigo = WVector(Da)
        XDate = Date$
        
        Call calcula_datos
        
        Call Calcula_Costo(WCodigo, Costo)
        WCosto = Str$(Costo)
        WStock = Str$(WEntradas - WSalidas)
        
        XParam = "'" + XCodigo + "','" _
                + WStock + "','" _
                + WCosto + "'"
                                           
        spTerminado = "ModificaTerminadoStock " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Da
    
    Listado.WindowTitle = "Listado de Valorizacion de Producto Terminado a Fecha"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaEnv.Envase} in " + DesdeEnv.Text + " to " + HastaEnv.Text
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Terminado.Codigo, Terminado.Descripcion, Terminado.Costo, Terminado.Stock " _
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

    Rem PROCESA LAS ESTADISTICAS
    
    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
        rstTerminado.Close
    End If
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spEstadistica = "ListaEstadisticaRepro" + XParam
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
                
                WAno = Right$(rstEstadistica!Fecha, 4)
                WMes = Mid$(rstEstadistica!Fecha, 4, 2)
                WDia = Left$(rstEstadistica!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                
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
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro1" + XParam
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
                    
                WAno = Right$(rstHoja!Fecha, 4)
                WMes = Mid$(rstHoja!Fecha, 4, 2)
                WDia = Left$(rstHoja!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                    
                    If rstHoja!Tipo = "T" Then
                        WSalidas = WSalidas + rstHoja!Cantidad
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
    End If
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro2" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                WAno = Right$(rstHoja!Fecha, 4)
                WMes = Mid$(rstHoja!Fecha, 4, 2)
                WDia = Left$(rstHoja!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                
                    If Val(rstHoja!Renglon) = 1 And rstHoja!Real <> 0 Then
                        WEntradas = WEntradas + rstHoja!Real
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
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovvar = "ListaMovvarRepro" + XParam
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
                        
                WAno = Right$(rstMovvar!Fecha, 4)
                WMes = Mid$(rstMovvar!Fecha, 4, 2)
                WDia = Left$(rstMovvar!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                
                If rstMovvar!Tipo = "T" Then
                
                    If rstMovvar!Movi = "E" Then
                        WEntradas = WEntradas + rstMovvar!Cantidad
                            Else
                        WSalidas = WSalidas + rstMovvar!Cantidad
                    End If
                
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
                 + WTerminado + "'"
    spMovguia = "ListaMovguiaRepro" + XParam
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
                        
                WAno = Right$(rstMovguia!Fecha, 4)
                WMes = Mid$(rstMovguia!Fecha, 4, 2)
                WDia = Left$(rstMovguia!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                
                If rstMovguia!Tipo = "T" Then
                
                    If rstMovguia!Movi = "E" Then
                        WEntradas = WEntradas + rstMovguia!Cantidad
                            Else
                        WSalidas = WSalidas + rstMovguia!Cantidad
                    End If
                
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
                 + WTerminado + "'"
    spMovlab = "ListaMovlabRepro" + XParam
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
                
                WAno = Right$(rstMovlab!Fecha, 4)
                WMes = Mid$(rstMovlab!Fecha, 4, 2)
                WDia = Left$(rstMovlab!Fecha, 2)
                WCompara = WAno + WMes + WDia
                       
                If WCompara <= WFechaord Then
                
                If rstMovlab!Tipo = "T" Then
                
                    If rstMovlab!Movi = "E" Then
                        WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                        WSalidas = WSalidas + rstMovlab!Cantidad
                    End If
                
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
                
                If rstConsig!Marca <> "X" Then
                
                    WAno = Right$(rstConsig!Fecha, 4)
                    WMes = Mid$(rstConsig!Fecha, 4, 2)
                    WDia = Left$(rstConsig!Fecha, 2)
                    WCompara = WAno + WMes + WDia
                       
                    If WCompara <= WFechaord Then
                        WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                        WSalidas = WSalidas + WCantidad
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
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Private Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgStock2.Caption = "Listado de Valorizacion de Producto Terminado a Fecha :  " + !Nombre
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
    
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
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
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub


