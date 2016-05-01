VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgValcar 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Valores en Cartera"
   ClientHeight    =   7380
   ClientLeft      =   3150
   ClientTop       =   735
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   ScaleHeight     =   7380
   ScaleWidth      =   5400
   Begin VB.Frame Frame2 
      Height          =   5175
      Left            =   480
      TabIndex        =   8
      Top             =   120
      Width           =   3855
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
         Height          =   495
         Left            =   2640
         TabIndex        =   15
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Cliente 
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
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   14
         Text            =   " "
         Top             =   2880
         Width           =   975
      End
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence2 
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
      Begin MSMask.MaskEdBox Vence1 
         Height          =   375
         Left            =   1080
         TabIndex        =   0
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Left            =   600
         TabIndex        =   11
         Top             =   4680
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
         Left            =   2040
         TabIndex        =   10
         Top             =   4680
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
         Left            =   240
         TabIndex        =   9
         Top             =   4080
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
         Left            =   1440
         TabIndex        =   4
         Top             =   4080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   255
         Left            =   1920
         TabIndex        =   16
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   255
         Left            =   1920
         TabIndex        =   17
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         Left            =   480
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente"
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
         Left            =   480
         TabIndex        =   13
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Parametros de Fechas"
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
         Left            =   840
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "valcar.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Saldos de Cuenta Corriente de Proveedores"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   2640
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
      Height          =   1815
      ItemData        =   "valcar.frx":0000
      Left            =   480
      List            =   "valcar.frx":0007
      TabIndex        =   6
      Top             =   5400
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgValcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WAuxiliar As String
Private WLinea As Single
Private Cheques(10) As Double
Private Impre(10) As Double
Private WTotal(10) As Double
Private WRecibo As Double
Private WCheque As String
Private WBanco As String
Private Impre1 As String
Private Impre2 As String
Private Impre3 As String
Private Impre4 As String
Private Impre5 As String
Private Impre6 As String
Private da As Single
Private WCliente As String
Dim rstRecibo As Recordset
Dim spRecibo As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String

Private Sub Acepta_Click()

    On Error GoTo Control_Error
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Actividad = "."
            .Update
        End If
    End With
    
    With rstValcar
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

    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)
    WDesdefec = Right$(Desdefec.Text, 4) + Mid$(Desdefec.Text, 4, 2) + Left$(Desdefec.Text, 2)
    WHastafec = Right$(Hastafec.Text, 4) + Mid$(Hastafec.Text, 4, 2) + Left$(Hastafec.Text, 2)
    
    Erase WTotal
    
    spRecibos = "ListaRecibosCartera"
    Set rstRecibos = db.OpenRecordset(spRecibos, dbOpenSnapshot, dbSQLPassThrough)
    If rstRecibos.RecordCount > 0 Then

    With rstRecibos
        .MoveFirst
        Do
        
            If .EOF = False Then
            
                If Val(!Tiporeg) = 2 Then
                
                    If Val(!Tipo2) = 2 And !Estado2 <> "X" Then
                    
                        WSaldo = !Importe2
                        Wvencimiento = Right$(!Fecha2, 4) + Mid$(!Fecha2, 4, 2) + Left$(!Fecha2, 2)
                        WCliente = !Cliente
                    
                        Erase Impre
                    
                        If WDesdefec <= Wvencimiento And WHastafec >= Wvencimiento Then
                        
                            If Cliente.Text = "" Or Cliente.Text = WCliente Then
                            
                                If Wvencimiento <= Fecha1 Then
                                    Impre(1) = WSaldo
                                        Else
                                    If Wvencimiento > Fecha1 And Wvencimiento <= Fecha2 Then
                                        Impre(2) = WSaldo
                                            Else
                                        If Wvencimiento > Fecha2 And Wvencimiento <= Fecha3 Then
                                            Impre(3) = WSaldo
                                                Else
                                            If Wvencimiento > Fecha3 And Wvencimiento <= Fecha4 Then
                                                Impre(4) = WSaldo
                                                    Else
                                                Impre(5) = WSaldo
                                            End If
                                        End If
                                    End If
                                End If
                                
                                If Impre(5) = 0 Then
                        
                                    WRecibo = !Recibo
                                    WCheque = !Numero2
                                    WBanco = !Banco2
                            
                                    Rem spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                                    Rem Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                                    Rem If rstCliente.RecordCount > 0 Then
                                    Rem     WRazon = rstCliente!Razon
                                    Rem     rstCliente.Clone
                                    Rem         Else
                                    Rem     WRazon = "."
                                    Rem End If
                        
                                    With rstValcar
                                        .Index = "Clave"
                                        .AddNew
                                        !Recibo = WRecibo
                                        !Cliente = WCliente
                                        !Cheque = WCheque
                                        !Banco = WBanco
                                        !Impo1 = Impre(1)
                                        !Impo2 = Impre(2)
                                        !Impo3 = Impre(3)
                                        !Impo4 = Impre(4)
                                        !Impo5 = Impre(5)
                                        !Razon = WRazon
                                        !Titulo1 = Vence1.Text
                                        !Titulo2 = Vence2.Text
                                        !Titulo3 = Vence3.Text
                                        !Titulo4 = Vence4.Text
                                        !Titulo5 = "Posterior"
                                        .Update
                                    End With
                                    
                                End If
                                
                            End If
                            
                        End If
                        
                    End If
                    
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    End If
    
    Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    Listado.WindowTitle = "Listado de Valores en cartera"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    Rem Listado.GroupSelectionFormula = "{Cliente.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Cliente.Text + Chr$(34)
    If Impresora.Value = True Then
       Listado.Destination = 1
           Else
       Listado.Destination = 0
    End If
    Listado.Action = 1
    Exit Sub
    
Control_Error:
     coderr = Err
     Resume Next
     Call Errores(coderr, "Cuenta", "No existe registro en el archivo")
    
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    Vence1.SetFocus
    PrgValcar.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Valcar
End Sub

Private Sub Vence1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Or Vence1.Text = "  /  /    " Or Vence1.Text = "00/00/0000" Then
            Vence2.SetFocus
                Else
            Vence1.SetFocus
        End If
    End If
End Sub

Private Sub Vence2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence2.Text, Auxi)
        If Auxi = "S" Or Vence2.Text = "  /  /    " Or Vence2.Text = "00/00/0000" Then
            Vence3.SetFocus
                Else
            Vence2.SetFocus
        End If
    End If
End Sub

Private Sub Vence3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence3.Text, Auxi)
        If Auxi = "S" Or Vence3.Text = "  /  /    " Or Vence3.Text = "00/00/0000" Then
            Vence4.SetFocus
                Else
            Vence3.SetFocus
        End If
    End If
End Sub

Private Sub Vence4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence4.Text, Auxi)
        If Auxi = "S" Or Vence4.Text = "  /  /    " Or Vence4.Text = "00/00/0000" Then
            Cliente.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
End Sub

Private Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desdefec.Text, Auxi)
        If Auxi = "S" Then
            Hastafec.SetFocus
                Else
            Desdefec.SetFocus
        End If
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hastafec.Text, Auxi)
        If Auxi = "S" Then
            Vence1.SetFocus
                Else
            Hastafec.SetFocus
        End If
    End If
End Sub


Sub Form_Load()

    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    Cliente.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spCliente = "ListaClientes"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        With rstCliente
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstCliente!Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCliente.Close
    End If
            
    Pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    WCliente = WIndice.List(Indice)
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        Cliente.Text = rstCliente!Cliente
        rstCliente.Close
                Else
        Cliente.Text = WCliente
    End If
    Cliente.SetFocus
End Sub



