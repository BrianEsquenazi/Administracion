VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListPago 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Ordenes de Pago"
   ClientHeight    =   4785
   ClientLeft      =   2925
   ClientTop       =   2415
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2175
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdefecha 
         Height          =   300
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2040
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   3480
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3480
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Desde fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5160
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wlistord.rpt"
      Destination     =   1
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
      Left            =   4680
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "listpago.frx":0000
      Left            =   0
      List            =   "listpago.frx":0007
      TabIndex        =   3
      Top             =   3240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4920
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPagos As Recordset
Dim spPagos As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo Control_Error

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !varios = "Desde el " + Desdefecha.Text + " hasta el " + HastaFecha.Text
            .Update
        End If
    End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstListado1
        .Index = "Clave"
        .MoveFirst
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
            
    spPagos = "ListaPagos"
    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
    If rstPagos.RecordCount > 0 Then
            
        With rstPagos
            .MoveFirst
            Do
            
                If WDesde <= !FechaOrd And !FechaOrd <= WHasta Then
                
                If !Importe1 <> 0 Then
                
                    WOrden = !Orden
                    WRenglon = !Renglon
                    WProveedor = !Proveedor
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WImporte = !Importe
                    WRetencion = !Retencion
                    WObservaciones = !Observaciones
                    WCuenta = !Cuenta
                    WTipoord = !TipoOrd
                    WTiporeg = !Tiporeg
                    WTipo1 = !Tipo1
                    WLetra1 = !Letra1
                    WPunto1 = !Punto1
                    WNumero1 = !Numero1
                    WImporte1 = !Importe1
                    WObservaciones2 = !Observaciones2
                    WTipo2 = !Tipo2
                    WNumero2 = !Numero2
                    WFecha2 = !Fecha2
                    WFechaord2 = !FechaOrd2
                    WBanco2 = !Banco2
                    WImporte2 = !Importe2
                    WClave = !Clave
                            
                    With rstListado1
                        .Index = "Clave"
                        .AddNew
                        !Orden = WOrden
                        !Renglon = WRenglon
                        !Proveedor = WProveedor
                        !Fecha = WFecha
                        !FechaOrd = WFechaord
                        !Importe = WImporte
                        !Retencion = WRetencion
                        !Observaciones = WObservaciones
                        !Cuenta = WCuenta
                        !TipoOrd = WTipoord
                        !Tiporeg = WTiporeg
                        !Tipo1 = WTipo1
                        !Letra1 = WLetra1
                        !Punto1 = WPunto1
                        !Numero1 = WNumero1
                        !Importe1 = WImporte1
                        !Observaciones2 = WObservaciones2
                        !Tipo2 = WTipo2
                        !Numero2 = WNumero2
                        !Fecha2 = WFecha2
                        !FechaOrd2 = WFechaord2
                        !Banco2 = WBanco
                        !Importe2 = WImporte2
                        !Empresa = WEmpresa
                        !Clave = WClave
                        !Titulo = WTitulo
                        !Titulo1 = "Desde el " + Desdefecha.Text + " hasta el " + HastaFecha.Text
                        .Update
                    End With
                End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    End If
    
    With rstListado1
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
                .Edit
                
                WProveedor = !Proveedor
                WNombre = ""
                
                spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                If RstProveedor.RecordCount > 0 Then
                    WNombre = RstProveedor!Nombre
                    RstProveedor.Close
                End If
                
                !Nombre = WNombre
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    

    Listado.WindowTitle = "Listado de Ordenes de Pago"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Uno = "{Listado1.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.GroupSelectionFormula = Uno
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub
    
Control_Error:
     coderr = Err
     Resume Next
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
    
End Sub

Private Sub Cancela_click()

    With rstAuxiliar
        .Close
    End With
    With rstListado1
        .Close
    End With
    Desdefecha.SetFocus
    PrgListPago.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Listado1
End Sub


Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hastafecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefecha.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Sub Form_Load()
    Rem  Desde.Text = ""
    Rem Hasta.Text = ""
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub
