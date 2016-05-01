VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgControl 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Control de Ordenes de Compra"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1815
      Left            =   840
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin VB.TextBox Hasta 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Desde 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Orden"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Orden"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6120
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WControl.rpt"
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
      Left            =   5760
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "control.frx":0000
      Left            =   840
      List            =   "control.frx":0007
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim Auxiliar(10000, 5)
Dim XParam As String

Private Sub Acepta_Click()

    Da = 0
    With rstControl
        .Index = "Orden"
        .Seek ">=", 0
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
    
    Erase Auxiliar
    Renglon = 0
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
                            
    spOrden = "ListaOrdenDesdeHasta " + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
        If rstOrden.RecordCount > 0 Then
    
            With rstOrden
    
                .MoveFirst
            
                Do
                
                    WOrden = rstOrden!Orden
                    WFecha = rstOrden!Fecha
                    WProveedor = rstOrden!Proveedor
                    WArticulo = rstOrden!Articulo
                    WCantidad = rstOrden!Cantidad
                
                    Renglon = Renglon + 1
                
                    Auxiliar(Renglon, 1) = WOrden
                    Auxiliar(Renglon, 2) = WFecha
                    Auxiliar(Renglon, 3) = WProveedor
                    Auxiliar(Renglon, 4) = WArticulo
                    Auxiliar(Renglon, 5) = WCantidad
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
        End With
    End If
    
    Graba = "S"

    For Da = 1 To Renglon
    
        WOrden = Auxiliar(Da, 1)
        WFecha = Auxiliar(Da, 2)
        WProveedor = Auxiliar(Da, 3)
        WArticulo = Auxiliar(Da, 4)
        WCantidad = Auxiliar(Da, 5)
        
        WIngre = "S"
                
        spLaudo = "ListaLaudoOrden " + "'" + Str$(WOrden) + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstLaudo.RecordCount > 0 Then
                
            With rstLaudo
    
                .MoveFirst
            
                Do
                    
                    If WArticulo = rstLaudo!Articulo Then
            
                        WLaudo = rstLaudo!Laudo
                        WRechazo = rstLaudo!Rechazo
                        WFecha1 = rstLaudo!Fecha
                        WLiberadaAnt = IIf(IsNull(rstLaudo!Liberadaant), "0", rstLaudo!Liberadaant)
                        WDevueltaAnt = IIf(IsNull(rstLaudo!devueltaant), "0", rstLaudo!devueltaant)
                        If WLiberadaAnt <> 0 Then
                            WCantidad1 = WLiberadaAnt
                                Else
                            WCantidad1 = rstLaudo!Liberada
                        End If
                        If WDevueltaAnt <> 0 Then
                            WCantidad2 = WDevueltaAnt
                                Else
                            WCantidad2 = rstLaudo!devuelta
                        End If
                
                        If WCantidad1 <> 0 Then
                        
                            With rstControl
                                .AddNew
                                !Orden = WOrden
                                !Fecha = WFecha
                                !Proveedor = WProveedor
                                !Articulo = WArticulo
                                !Cantidad = WCantidad
                                !Comprobante = "Laudo " + Str$(WLaudo)
                                !Fecha1 = WFecha1
                                !Cantidad1 = WCantidad1 * -1
                                !Cantidad2 = 0
                                .Update
                                WIngre = "N"
                            End With
                            Graba = "N"
                            
                        End If
                        
                        If WCantidad2 <> 0 Then
                            With rstControl
                                .AddNew
                                !Orden = WOrden
                                !Fecha = WFecha
                                !Proveedor = WProveedor
                                !Articulo = WArticulo
                                !Cantidad = WCantidad
                                !Comprobante = "Rechazo " + Str$(WRechazo)
                                !Fecha1 = WFecha1
                                !Cantidad2 = WCantidad2 * -1
                                !Cantidad1 = 0
                                .Update
                                WIngre = "N"
                            End With
                            Graba = "N"
                        End If
                        
                    End If
            
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
                
            End With
        End If
                
                
        spInforme = "ListaInformeOrden " + "'" + Str$(WOrden) + "'"
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
        If rstInforme.RecordCount > 0 Then
                
            With rstInforme
    
                .MoveFirst
            
                Do
                    
                    If WArticulo = rstInforme!Articulo Then
                        
                        WInforme = rstInforme!Informe
                        WFecha1 = rstInforme!Fecha
                        WCantidad1 = rstInforme!Cantidad
                
                        With rstControl
                            .AddNew
                            !Orden = WOrden
                            !Fecha = WFecha
                            !Proveedor = WProveedor
                            !Articulo = WArticulo
                            !Cantidad = WCantidad
                            !Comprobante = "Informe " + Str$(WInforme)
                            !Fecha1 = WFecha1
                            !Cantidad1 = WCantidad1
                            !Cantidad2 = 0
                            .Update
                            WIngre = "N"
                        End With
                        Graba = "N"
                        
                    End If
            
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
                
            End With
                
        End If
        
        If WIngre = "S" Then
            With rstControl
                .AddNew
                !Orden = WOrden
                !Fecha = WFecha
                !Proveedor = WProveedor
                !Articulo = WArticulo
                !Cantidad = WCantidad
                !Comprobante = ""
                !Fecha1 = ""
                !Cantidad1 = 0
                !Cantidad2 = 0
                .Update
            End With
            Graba = "N"
        End If
        
    Next Da
    
    If Graba = "S" Then
        With rstControl
            .AddNew
            !Orden = WOrden
            !Fecha = WFecha
            !Proveedor = WProveedor
            !Articulo = WArticulo
            !Cantidad = WCantidad
            !Comprobante = ""
            !Fecha1 = ""
            !Cantidad2 = 0
            !Cantidad1 = 0
            .Update
            WIngre = "N"
        End With
    End If
    
    Da = 0
    With rstControl
        .Index = "Orden"
        .Seek ">=", 0
        If .NoMatch = False Then
            Do
                .Edit
                
                WDescriArticulo = ""
                WDescriProvdeedor = ""
                WArticulo = !Articulo
                WProveedor = !Proveedor
                
                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    WDescriArticulo = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WDescriProveedor = RstProveedor!Nombre
                End If
                
                !DescriArticulo = WDescriArticulo
                !DescriProveedor = WDescriProveedor
                
                .Update
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Control de Ordenes de Compra"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Listado.GroupSelectionFormula = "{Control.Orden} in " + Desde.Text + " to " + Hasta.Text
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    
    Desde.SetFocus
    PrgControl.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
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
    OPEN_FILE_Control
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgControl.Caption = "Listado de Control de Ordenes de Compra :  " + !Nombre
        End If
    End With

    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = True
End Sub


