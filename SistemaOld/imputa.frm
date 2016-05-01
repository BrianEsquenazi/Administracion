VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImputa 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Imputaciones Contables"
   ClientHeight    =   5880
   ClientLeft      =   3165
   ClientTop       =   1620
   ClientWidth     =   5670
   LinkTopic       =   "Form2"
   ScaleHeight     =   5880
   ScaleWidth      =   5670
   Begin VB.ListBox Pantalla 
      Height          =   1425
      Left            =   360
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4440
      TabIndex        =   15
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4440
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3135
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Text            =   " "
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox HastaCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   3
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox DesdeCuenta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   2
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   282
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   2520
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Desde Cuenta"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "09"
      Destination     =   1
      WindowTitle     =   "Listado de Imputaciones Contables"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgImputa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstImpu As Recordset
Dim spImpu As String
Dim rstCuenta As Recordset
Dim spCuenta As String
Dim XParam As String


Private Sub Acepta_Click()

    Rem On Error GoTo WError
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Proveedor SET "
    ZSql = ZSql + " Impre = " + "'" + "" + "'"
    spProveedor = ZSql
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    
    listado.WindowTitle = "Listado de Imputaciones Contables de Compras"
    listado.WindowTop = 0
    listado.WindowLeft = 0
    listado.WindowWidth = Screen.Width
    listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia
    
    Rem If Tipo1.Value = True Then
    Rem     WTipo = "1"
    Rem End If
    Rem If Tipo2.Value = True Then
    Rem     WTipo = "2"
    Rem End If
    Rem If Tipo3.Value = True Then
    Rem     WTipo = "3"
    Rem End If
    Rem If Tipo4.Value = True Then
    Rem     WTipo = "4"
    Rem End If
    
    WTipo = 2
    
    Uno = ""
    Dos = "{Imputac.Cuenta} in " + Chr$(34) + DesdeCuenta.Text + Chr$(34) + " to " + Chr$(34) + HastaCuenta.Text + Chr$(34)
    Tres = "{Imputac.Tipomovi} in " + Chr$(34) + WTipo + Chr$(34) + " to " + Chr$(34) + WTipo + Chr$(34)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WTitulo = !Nombre
        End If
    End With
    
    da = ""
    With rstImputac
        .Index = "Clave"
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
    
    Rem XParam = "'" + WDesde + "','" _
    rem              + WHasta + "','" _
    rem              + DesdeCuenta.Text + "','" _
    rem              + HastaCuenta.Text + "'"
    Rem spImpu = "ListaImputacDesdeHasta " + XParam
    
    ZSql = ""
    ZSql = ZSql + "Select Imputac.TipoMovi, Imputac.NroInterno, Imputac.Proveedor, Imputac.TipoComp, Imputac.LetraComp, Imputac.PuntoComp, Imputac.NroComp, Imputac.Renglon, Imputac.Fecha, Imputac.Observaciones, Imputac.Cuenta, Imputac.Debito, Imputac.Credito, Imputac.fechaord, Imputac.Titulo, Imputac.Empresa, Imputac.Clave, Cuenta.Cuenta, Cuenta.Descripcion, Ivacomp.Periodo"
    ZSql = ZSql + " FROM Imputac, Cuenta, IvaComp"
    ZSql = ZSql + " Where Imputac.Cuenta >= " + "'" + DesdeCuenta.Text + "'"
    ZSql = ZSql + " and Imputac.Cuenta <= " + "'" + HastaCuenta.Text + "'"
    ZSql = ZSql + " and Imputac.Cuenta = Cuenta.Cuenta"
    ZSql = ZSql + " and Imputac.NroInterno = IvaComp.NroInterno"
    spImpu = ZSql
    Set RstImpu = db.OpenRecordset(spImpu, dbOpenSnapshot, dbSQLPassThrough)
    If RstImpu.RecordCount > 0 Then
    
    With RstImpu
    
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                WFecha = Right$(!Periodo, 4) + Mid$(!Periodo, 4, 2) + Left$(!Periodo, 2)
                If WDesde <= WFecha And WFecha <= WHasta Then
            
                XTipomovi = !Tipomovi
                XNroInterno = !NroInterno
                XProveedor = !Proveedor
                XTipocomp = !TipoComp
                XLetracomp = !LetraComp
                XPuntocomp = !PuntoComp
                XNrocomp = !NroComp
                XRenglon = !Renglon
                XFecha = !Fecha
                XObservaciones = !Observaciones
                XCuenta = !Cuenta
                XDebito = !Debito
                XCredito = !Credito
                XFechaOrd = !FechaOrd
                XTitulo = !Titulo
                XEmpresa = !Empresa
                XClave = !Clave
                XNombre = !Descripcion
                
                With rstImputac
                
                    .Index = "Clave"
                    .Seek "=", XClave
                    If .NoMatch Then
                        .AddNew
                        !Tipomovi = XTipomovi
                        !NroInterno = XNroInterno
                        !Proveedor = XProveedor
                        !TipoComp = XTipocomp
                        !LetraComp = XLetracomp
                        !PuntoComp = XPuntocomp
                        !NroComp = XNrocomp
                        !Renglon = XRenglon
                        !Fecha = XFecha
                        !Observaciones = XObservaciones
                        !Cuenta = XCuenta
                        !Debito = XDebito
                        !Credito = XCredito
                        !FechaOrd = XFechaOrd
                        !Titulo = XTitulo
                        !Empresa = XEmpresa
                        !Clave = XClave
                        !Titulolist = WTitulo
                        !debitoList = XDebito
                        !CreditoList = XCredito
                        !Nombre = XNombre
                        .Update
                        .Bookmark = .LastModified
                    End If
                End With
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        End If
        
    End With
    RstImpu.Close
    
    
    End If
    
    
    If Val(Wempresa) = 8 Then
     
        With rstImputac
            .Index = "Clave"
            .MoveFirst
            Do
            
                If !Cuenta = "2001" Then
                
                    WNroInterno = !NroInterno
                    WContado = 0
                    
                    spIvaComp = "ConsultaIvacomp " + "'" + Str$(WNroInterno) + "'"
                    Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                    If rstIvaComp.RecordCount > 0 Then
                        WContado = rstIvaComp!Contado
                        rstIvaComp.Close
                    End If
                    
                    If Val(WContado) = 3 Then
                        .Edit
                        !Cuenta = "2046"
                        !Nombre = "PYME NACION"
                        .Update
                    End If
                    
                End If
    
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
    End If
    
    
    
    
    
    
    
    
    
    
    Rem With rstImputac
    Rem         .Index = "Clave"
    Rem         .MoveFirst
    Rem         Do
    Rem             .Edit
    Rem             !debitoList = 0
    Rem             !CreditoList = 0
    Rem             If !FechaOrd >= WDesde And !FechaOrd <= WHasta Then
    Rem                 If Val(!Cuenta) >= Val(DesdeCuenta.Text) And Val(!Cuenta) <= Val(HastaCuenta.Text) Then
    Rem                     Rem If !Tipomovi = WTipo Then
    Rem                         !debitoList = !Debito
    Rem                         !CreditoList = !Credito
    Rem                     Rem End If
    Rem                 End If
    Rem             End If
    Rem
    Rem             WNombre = ""
    Rem             WCuenta = !Cuenta
    Rem
    Rem             spCuenta = "Consultacuentas " + "'" + WCuenta + "'"
    Rem             Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstCuenta.RecordCount > 0 Then
    Rem                 WNombre = rstCuenta!Descripcion
    Rem                 rstCuenta.Close
    Rem             End If
    Rem
    Rem             !Nombre = WNombre
    Rem
    Rem             .Update
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    
    
    
    Rem With rstImputac
    Rem         .Index = "Clave"
   Rem          .MoveFirst
    Rem         Do
    Rem
    Rem             ZProveedor = !Proveedor
    Rem
    Rem             ZSql = ""
    Rem             ZSql = ZSql + "UPDATE Proveedor SET "
    Rem             ZSql = ZSql + " Impre = " + "'" + "S" + "'"
    Rem             ZSql = ZSql + " Where Proveedor = " + "'" + ZProveedor + "'"
     Rem            spProveedor = ZSql
    Rem             Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    
    
    
    
    
    
    
    
    
    listado.GroupSelectionFormula = Dos + " and " + Tres
    listado.SelectionFormula = Dos + " and " + Tres
    
    If Tipo.ListIndex = 0 Then
        listado.ReportFileName = "WImputa.rpt"
            Else
        listado.ReportFileName = "WImputa2.rpt"
    End If
        
    If Impresora.Value = True Then
        listado.Destination = 1
            Else
        listado.Destination = 0
    End If
    
    listado.DataFiles(0) = Wempresa + "auxi.mdb"
    
    listado.Action = 1
    
    Exit Sub
    
WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstImputac
        .Close
    End With
    Desde.SetFocus
    PrgImputa.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.Text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Imputac
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            DesdeCuenta.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCuenta.SetFocus
    End If
End Sub

Private Sub HastaCuenta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    spCuenta = "ListaCuentas"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
            
    With rstCuenta
        .MoveFirst
        Do
            If .EOF = False Then
                IngresaItem = rstCuenta!Cuenta + " " + rstCuenta!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstCuenta!Cuenta
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstCuenta.Close
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WCuenta = WIndice.List(Indice)
    spCuenta = "ConsultaCuentas " + "'" + WCuenta + "'"
    Set rstCuenta = db.OpenRecordset(spCuenta, dbOpenSnapshot, dbSQLPassThrough)
    If rstCuenta.RecordCount > 0 Then
        DesdeCuenta.Text = rstCuenta!Cuenta
        HastaCuenta.Text = rstCuenta!Cuenta
        rstCuenta.Close
                Else
        DesdeCuenta.Text = WCuenta
        HastaCuenta.Text = WCuenta
    End If
    DesdeCuenta.SetFocus
    
End Sub

Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeCuenta.Text = " "
    HastaCuenta.Text = "999999999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
    Tipo.Clear
    
    Tipo.AddItem "Completo"
    Tipo.AddItem "Resumido"
    
    Tipo.ListIndex = 0
    
End Sub

