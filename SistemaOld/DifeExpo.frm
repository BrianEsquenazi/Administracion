VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgDifeExpo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Diferencias de Cambio (Cobranzas)"
   ClientHeight    =   4785
   ClientLeft      =   2925
   ClientTop       =   2415
   ClientWidth     =   6240
   LinkTopic       =   "Form2"
   ScaleHeight     =   4785
   ScaleWidth      =   6240
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   2295
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
         Left            =   2280
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   1080
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
      ReportFileName  =   "WDifeI.rpt"
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
      ItemData        =   "DifeExpo.frx":0000
      Left            =   0
      List            =   "DifeExpo.frx":0007
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
Attribute VB_Name = "PrgDifeExpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim rstCambios As Recordset
Dim spCambios As String
Dim XParam As String
Dim Vector(10000, 4) As String
Dim WClave As String
Dim WFecha As String
Dim WTipo As String
Dim WNumero As String
Dim Paridad1 As String
Dim Paridad2 As String
Dim WFechaFactura As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Difefrencias de Cambio (Cobranzas)"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WAno = Right$(Desdefecha.Text, 4)
    WMes = Mid$(Desdefecha.Text, 4, 2)
    WDia = Left$(Desdefecha.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(HastaFecha.Text, 4)
    WMes = Mid$(HastaFecha.Text, 4, 2)
    WDia = Left$(HastaFecha.Text, 2)
    WHasta = WAno + WMes + WDia

    spRecibo = "ModificaReciboImpolista0"
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    
    XParam = "'" + WDesde + "','" _
                 + WHasta + "'"
    spRecibo = "ListaRecibosDifeI" + XParam
    Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibo.RecordCount > 0 Then
        With rstRecibo
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    Vector(Renglon, 1) = rstRecibo!Clave
                    Vector(Renglon, 2) = rstRecibo!Fecha
                    Vector(Renglon, 3) = rstRecibo!Tipo1
                    Vector(Renglon, 4) = rstRecibo!Numero1
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstRecibo.Close
    End If
    
    For Cicla = 1 To Renglon
    
        WClave = Vector(Cicla, 1)
        WFecha = Vector(Cicla, 2)
        WTipo = Vector(Cicla, 3)
        WNumero = Vector(Cicla, 4)
        
        spCambios = "ConsultaCambio  " + "'" + WFecha + "'"
        Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        If rstCambios.RecordCount > 0 Then
            Paridad2 = Str$(rstCambios!Cambio)
            rstCambios.Close
                Else
            Paridad2 = "1"
        End If
        
        ClaveCtacte = WTipo + WNumero + "01"
        spCtacte = "ConsultaCtacte " + "'" + ClaveCtacte + "'"
        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtacte.RecordCount > 0 Then
            Paridad1 = Str$(rstCtacte!Paridad)
            If Val(Paridad1) = 0 Then
                Paridad1 = "1"
            End If
            WFechaFactura = rstCtacte!Fecha
            rstCtacte.Close
                Else
            WFechaFactura = "00/00/0000"
            Paridad1 = "1"
        End If
        
        Rem spCambios = "ConsultaCambio  " + "'" + WFechaFactura + "'"
        Rem Set rstCambios = db.OpenRecordset(spCambios, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstCambios.RecordCount > 0 Then
        Rem     Paridad1 = Str$(rstCambios!Cambio)
        Rem     rstCambios.Close
        Rem         Else
        Rem     Paridad1 = "1"
        Rem End If
        
        XParam = "'" + WClave + "','" _
                    + Paridad1 + "','" _
                    + Paridad2 + "','" _
                    + WFechaFactura + "'"
        spRecibo = "ModificaReciboDifeI " + XParam
        Set rstRecibo = db.OpenRecordset(spRecibo, dbOpenSnapshot, dbSQLPassThrough)
    
    Next Cicla

    Uno = "{reciboS.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
    Listado.GroupSelectionFormula = Uno
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    Listado.SQLQuery = "SELECT Recibos.Recibo, Recibos.Cliente, Recibos.Fecha, Recibos.Fechaord, Recibos.Tipo1, Recibos.Numero1, Recibos.Importe1, Recibos.Impolist, Recibos.Impo1list, " _
                        + "Cliente.Razon " _
                        + "From " _
                        + DSQ + ".dbo.Recibos Recibos, " _
                        + DSQ + ".dbo.Cliente Cliente " _
                        + "Where " _
                        + "Recibos.Cliente = Cliente.Cliente AND " _
                        + "Recibos.Fechaord >= '" + WDesde + "' AND Recibos.Fechaord <= '" + WHasta + "' AND " _
                        + "Recibos.Impolist <> 0. AND " _
                        + "Recibos.Impo1list <> 0."
                        
    Listado.DataFiles(2) = WEmpresa + "Auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    Desdefecha.SetFocus
    PrgDifeI.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Desdefecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.Text = Desdefecha.Text
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
    Desdefecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

