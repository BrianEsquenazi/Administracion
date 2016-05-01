VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgProyCtaCteAnalitico 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Proyeccion de Cuentas Corrientes de Clientes"
   ClientHeight    =   7770
   ClientLeft      =   1485
   ClientTop       =   750
   ClientWidth     =   9120
   LinkTopic       =   "Form2"
   ScaleHeight     =   7770
   ScaleWidth      =   9120
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Text            =   " "
      Top             =   3120
      Width           =   7695
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   480
      TabIndex        =   9
      Top             =   240
      Width           =   7695
      Begin VB.ComboBox Tipo 
         Height          =   315
         Left            =   3600
         TabIndex        =   17
         Top             =   1560
         Width           =   2295
      End
      Begin VB.TextBox HastaClie 
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   4
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox DesdeClie 
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
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   3
         Text            =   " "
         Top             =   720
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4080
         TabIndex        =   14
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2400
         TabIndex        =   13
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
         Height          =   495
         Left            =   5520
         TabIndex        =   12
         Top             =   360
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
         Height          =   495
         Left            =   5520
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   285
         Left            =   3600
         TabIndex        =   0
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
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
         Caption         =   "Hasta Cliente"
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
         Left            =   1560
         TabIndex        =   16
         Top             =   1560
         Width           =   1695
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Cliente"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Cliente"
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
         Left            =   1560
         TabIndex        =   10
         Top             =   720
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   0
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wccprv.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
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
      Left            =   -480
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   3765
      ItemData        =   "ProyCtaCteAnalitico.frx":0000
      Left            =   480
      List            =   "ProyCtaCteAnalitico.frx":0007
      TabIndex        =   7
      Top             =   3600
      Width           =   7695
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   8160
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   8160
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgProyCtaCteAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Acumula As Double
Private Pasa As Single
Private WSaldo As Double
Dim RstCtaPrv As Recordset
Dim spCtaprv As String
Dim rstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String
Dim XSaldo As Double
Dim XTotal As Double
Dim XPagos As Double
Dim XDiferencia As Double

Dim WDesde As String
Dim WHasta As String
Dim WFecha As String
Dim ZDias As Integer
Dim ZDia As String
Dim ZMes As String
Dim ZAno As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub Acepta_Click()

    Listado.WindowTitle = "Listado de Cuenta Corriente de Clientes Analitico"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WFecha = Fecha.Text
    Cicla = 0
    Do
    
        Cicla = Cicla + 1
        If Cicla = 1000 Then
          Rem Exit Sub
         Stop
       End If
        ZDias = DateDiff("d", WFecha, Fecha.Text)
        If ZDias = 30 Or ZDias = 31 Then
            WFecha1 = WFecha
            Exit Do
        End If
        
        ZDia = Mid$(WFecha, 1, 2)
        ZMes = Mid$(WFecha, 4, 2)
        ZAno = Mid$(WFecha, 7, 4)
        
        ZDia = Str$(Val(ZDia) - 1)
        If Val(ZDia) = 0 Then
            ZMes = Str$(Val(ZMes) - 1)
            If Val(ZMes) = 0 Then
                ZAno = Str$(Val(ZAno) - 1)
                ZMes = "12"
            End If
            If Val(ZMes) = 2 Then
                ZDia = "28"
                    Else
                ZDia = "30"
            End If
        End If
        
        Call Ceros(ZDia, 2)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        WFecha = ZDia + "/" + ZMes + "/" + ZAno
        
    Loop
    
    WFecha = Fecha.Text
    Cicla = 0
    Do
    
        Cicla = Cicla + 1
        If Cicla = 1000 Then Exit Sub

        ZDias = DateDiff("d", WFecha, Fecha.Text)
        If ZDias = 60 Or ZDias = 61 Then
            WFecha2 = WFecha
            Exit Do
        End If
        
        ZDia = Mid$(WFecha, 1, 2)
        ZMes = Mid$(WFecha, 4, 2)
        ZAno = Mid$(WFecha, 7, 4)
        
        ZDia = Str$(Val(ZDia) - 1)
        If Val(ZDia) = 0 Then
            ZMes = Str$(Val(ZMes) - 1)
            If Val(ZMes) = 0 Then
                ZAno = Str$(Val(ZAno) - 1)
                ZMes = "12"
            End If
            If Val(ZMes) = 2 Then
                ZDia = "28"
                    Else
                ZDia = "30"
            End If
        End If
        
        Call Ceros(ZDia, 2)
        Call Ceros(ZMes, 2)
        Call Ceros(ZAno, 4)
        
        WFecha = ZDia + "/" + ZMes + "/" + ZAno
        
    Loop
    
    WAno = Right$(WFecha1, 4)
    WMes = Mid$(WFecha1, 4, 2)
    WDia = Left$(WFecha1, 2)
    WFechaOrd1 = WAno + WMes + WDia
    
    WAno = Right$(WFecha2, 4)
    WMes = Mid$(WFecha2, 4, 2)
    WDia = Left$(WFecha2, 2)
    WFechaord2 = WAno + WMes + WDia
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Auxiliar SET "
    ZSql = ZSql + " Auxi1 = " + "'" + WFecha1 + "',"
    ZSql = ZSql + " Auxi2 = " + "'" + WFecha2 + "',"
    ZSql = ZSql + " Auxi3 = " + "'" + WFechaOrd1 + "',"
    ZSql = ZSql + " Auxi4 = " + "'" + WFechaord2 + "'"
    spAuxiliar = ZSql
    Set rstAuxiliar = db.OpenRecordset(spAuxiliar, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
    ZSql = ZSql + " Empresa = 1"
    spCtacte = ZSql
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
   
    If Tipo.ListIndex = 0 Then
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT CtaCte.Cliente, CtaCte.Saldo, CtaCte.OrdFecha, " _
                    + "Cliente.Razon, " _
                    + "Auxiliar.Nombre, Auxiliar.Auxi3, Auxiliar.Auxi4 " _
                    + "From " _
                    + DSQ + ".dbo.CtaCte CtaCte, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Auxiliar Auxiliar " _
                    + "Where " _
                    + "CtaCte.Cliente = Cliente.Cliente AND " _
                    + "CtaCte.Empresa = Auxiliar.Empresa AND " _
                    + "CtaCte.Cliente >= '" + DesdeClie.Text + "' AND " _
                    + "CtaCte.Cliente <= '" + HastaClie.Text + "' AND " _
                    + "(CtaCte.Saldo < -1 OR CtaCte.Saldo > 1)"
    
        Listado.Connect = Connect()
        
        Listado.ReportFileName = "WProyCtaCteAnalitico.rpt"
        
        Uno = "{CtaCte.Cliente} in " + Chr$(34) + DesdeClie.Text + Chr$(34) + " to " + Chr$(34) + HastaClie.Text + Chr$(34)
        Dos = " and ({CtaCte.Saldo} > 1 or  {CtaCte.Saldo} < -1)"
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
        
        Listado.Action = 1
        
            Else

        
        
       For i = 1 To 9
        
        Select Case i
          Case 1
             vend = "1"
          Case 2
             vend = "2"
          Case 3
             vend = "4"
          Case 4
             vend = "10"
         Case 5
             vend = "11"
          Case 6
             vend = "12"
         
         Case 7
             vend = "15"
         Case 8
             vend = "16"
          Case 9
             vend = "18"
        
        
        
        End Select
        
        DbConnect = db.Connect
        DSQ = getDatabase(DbConnect)
        Listado.SQLQuery = "SELECT CtaCte.Cliente, CtaCte.Saldo, CtaCte.OrdFecha, CtaCte.Paridad, CtaCte.Provincia, " _
                + "Cliente.Razon, Cliente.Vendedor, " _
                + "Auxiliar.Nombre, Auxiliar.Auxi3, Auxiliar.Auxi4, " _
                + "Vendedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.CtaCte CtaCte, " _
                + DSQ + ".dbo.Cliente Cliente, " _
                + DSQ + ".dbo.Auxiliar Auxiliar, " _
                + DSQ + ".dbo.Vendedor Vendedor " _
                + "Where " _
                + "CtaCte.Cliente = Cliente.Cliente AND " _
                + "CtaCte.Empresa = Auxiliar.Empresa AND " _
                + "Cliente.Vendedor = Vendedor.Vendedor AND " _
                + "Cliente.Vendedor = '" + vend + "' AND " _
                + "CtaCte.Cliente >= '" + DesdeClie.Text + "' AND " _
                + "CtaCte.Cliente <= '" + HastaClie.Text + "' AND " _
                + "(CtaCte.Saldo < -1 OR CtaCte.Saldo > 1)"
    
    
    Rem + "Cliente.Vendedor = Vendedor.Vendedor AND " _
        Listado.Connect = Connect()
        
        Listado.ReportFileName = "WProyCtaCteAnaliticoVende.rpt"
        
        Uno = "{CtaCte.Cliente} in " + Chr$(34) + DesdeClie.Text + Chr$(34) + " to " + Chr$(34) + HastaClie.Text + Chr$(34)
        Dos = " and ({CtaCte.Saldo} > 1 or  {CtaCte.Saldo} < -1)"
        
        Listado.GroupSelectionFormula = Uno + Dos
        Listado.SelectionFormula = Uno + Dos
        
        Listado.Action = 1
   Next i
    End If
            
End Sub

Private Sub Cancela_click()
    PrgProyCtaCteAnalitico.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub pantalla_Click()
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    DesdeClie.Text = WIndice.List(Indice)
    HastaClie.Text = WIndice.List(Indice)
End Sub

Sub Form_Load()

    Pantalla.Clear
    Ayuda.Text = ""
    
    Tipo.Clear
    
    Tipo.AddItem "Por Cliente"
    Tipo.AddItem "Por Vendedor"
    
    Tipo.ListIndex = 0
    

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    
    DesdeClie.Text = ""
    HastaClie.Text = ""
    
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

        Pantalla.Clear
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        If Ayuda.Text <> "" Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Razon LIKE " + "'" + "%" + Ayuda.Text + "%" + "'"
            ZSql = ZSql + " Order by Razon"
            spCliente = ZSql
                Else
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Order by Razon"
            spCliente = ZSql
        End If
    
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            With rstCliente
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cliente + "    " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCliente.Close
        End If
    
    End If

End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            DesdeClie.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub DesdeClie_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaClie.SetFocus
    End If
    If KeyAscii = 27 Then
        DesdeClie.Text = ""
    End If
End Sub

Private Sub HastaClie_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        HastaClie.Text = ""
    End If
End Sub








