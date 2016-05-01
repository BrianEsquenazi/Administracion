VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCash 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cash Flow"
   ClientHeight    =   6330
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6135
   LinkTopic       =   "Form2"
   ScaleHeight     =   6330
   ScaleWidth      =   6135
   Begin VB.Frame Frame4 
      Caption         =   "Vencimiento"
      Height          =   1335
      Left            =   4200
      TabIndex        =   25
      Top             =   4080
      Width           =   1815
      Begin VB.OptionButton Venci2 
         Caption         =   "Vencimiento 2"
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton Venci1 
         Caption         =   "Vencimiento 1"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Listado"
      Height          =   1335
      Left            =   4200
      TabIndex        =   19
      Top             =   2520
      Width           =   1935
      Begin VB.OptionButton Total 
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Documentos 
         Caption         =   "Documentos"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton CtaCte 
         Caption         =   "Cta.Cte."
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Moneda"
      Height          =   1215
      Left            =   4200
      TabIndex        =   18
      Top             =   1080
      Width           =   1815
      Begin VB.OptionButton Dolares 
         Caption         =   "Dolares"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Pesos 
         Caption         =   "Pesos"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   4335
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Vence4 
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence3 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence2 
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Vence1 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox Hasta 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox Desde 
         Height          =   285
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   3720
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3720
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
         TabIndex        =   8
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000A&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Parametros de Fechas"
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   3840
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCash.rpt"
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
      Left            =   4440
      TabIndex        =   4
      Top             =   5760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "CASH.frx":0000
      Left            =   600
      List            =   "CASH.frx":0007
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4320
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4320
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WSaldo As Double
Private Wvencimiento As String
Private WCliente As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim XParam As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Auxi1 = Vence1.Text
            !Auxi2 = Vence2.Text
            !Auxi3 = Vence3.Text
            !Auxi4 = Vence4.Text
            .Update
        End If
    End With

    Listado.WindowTitle = "Listado de Cash Flow"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Fecha1 = Right$(Vence1.Text, 4) + Mid$(Vence1.Text, 4, 2) + Left$(Vence1.Text, 2)
    Fecha2 = Right$(Vence2.Text, 4) + Mid$(Vence2.Text, 4, 2) + Left$(Vence2.Text, 2)
    Fecha3 = Right$(Vence3.Text, 4) + Mid$(Vence3.Text, 4, 2) + Left$(Vence3.Text, 2)
    Fecha4 = Right$(Vence4.Text, 4) + Mid$(Vence4.Text, 4, 2) + Left$(Vence4.Text, 2)
    
    DA = ""
    With rstCash
        .Index = "Cliente"
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
    
    spCtacte = "ListaCtacte"
    Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtacte.RecordCount > 0 Then
    
    With rstCtacte
            .MoveFirst
            Do
                If Pesos.Value = True Then
                    WSaldo = !Saldo
                            Else
                    WSaldo = !Saldous
                End If
                
                If WSaldo <> 0 Then
                
                If !Cliente >= Desde.Text And !Cliente <= Hasta.Text Then
                
                WPasa = "N"
                
                If CtaCte.Value = True Then
                    If !Tipo < 50 Then
                        WPasa = "S"
                    End If
                End If
    
                If Documentos.Value = True Then
                    If !Tipo >= 50 Then
                        WPasa = "S"
                    End If
                End If
    
                If Total.Value = True Then
                    WPasa = "S"
                End If
                
                If WPasa = "S" Then
                                    
                    If Venci1.Value = True Then
                        Wvencimiento = Right$(!Vencimiento, 4) + Mid$(!Vencimiento, 4, 2) + Left$(!Vencimiento, 2)
                            Else
                        Wvencimiento = Right$(!Vencimiento1, 4) + Mid$(!Vencimiento1, 4, 2) + Left$(!Vencimiento1, 2)
                    End If
                    WCliente = !Cliente
                    
                    With rstCash
                        .Index = "Cliente"
                        .Seek "=", WCliente
                        If .NoMatch = False Then
                            .Edit
                            !Importe6 = !Importe6 + WSaldo
                            If Wvencimiento <= Fecha1 Then
                                !Importe1 = !Importe1 + WSaldo
                                    Else
                                If Wvencimiento <= Fecha2 Then
                                    !Importe2 = !Importe2 + WSaldo
                                        Else
                                    If Wvencimiento <= Fecha3 Then
                                        !Importe3 = !Importe3 + WSaldo
                                            Else
                                        If Wvencimiento <= Fecha4 Then
                                            !Importe4 = !Importe4 + WSaldo
                                                Else
                                            !Importe5 = !Importe5 + WSaldo
                                        End If
                                    End If
                                End If
                            End If
                            .Update
                                Else
                            .AddNew
                            !Cliente = WCliente
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !Importe6 = !Importe6 + WSaldo
                            If Wvencimiento <= Fecha1 Then
                                !Importe1 = !Importe1 + WSaldo
                                    Else
                                If Wvencimiento <= Fecha2 Then
                                    !Importe2 = !Importe2 + WSaldo
                                        Else
                                    If Wvencimiento <= Fecha3 Then
                                        !Importe3 = !Importe3 + WSaldo
                                            Else
                                        If Wvencimiento <= Fecha4 Then
                                            !Importe4 = !Importe4 + WSaldo
                                                Else
                                            !Importe5 = !Importe5 + WSaldo
                                        End If
                                    End If
                                End If
                            End If
                            .Update
                        End If
                    End With
                End If
                End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    End If
    
    DA = ""
    With rstCash
        .Index = "Cliente"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                spCliente = "ConsultaCliente " + "'" + !Cliente + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WRazon = rstCliente!Razon
                    rstCliente.Close
                End If
                !Razon = WRazon
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    

    Listado.GroupSelectionFormula = "{Cash.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
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
    With rstAuxiliar
        .Close
    End With
    Desde.SetFocus
    PrgCash.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    spClientes = "ListaClienteConsulta"
    Set rstClientes = db.OpenRecordset(spClientes, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        With rstClientes
            .MoveFirst
            Do
                If .EOF = False Then
                    IngresaItem = rstClientes!Cliente + " " + rstClientes!Razon
                    Pantalla.AddItem IngresaItem
                    IngresaItem = rstClientes!Cliente
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstClientes.Close
    End If
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spClientes = "ConsultaClientes " + "'" + Claveven$ + "'"
    Set rstClientes = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstClientes.RecordCount > 0 Then
        Desde.Text = rstClientes!Cliente
        Hasta.Text = rstClientes!Cliente
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
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
        Vence1.SetFocus
    End If
End Sub

Private Sub Vence1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence1.Text, Auxi)
        If Auxi = "S" Then
            Vence2.SetFocus
                Else
            Vence1.SetFocus
        End If
    End If
End Sub

Private Sub Vence2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence2.Text, Auxi)
        If Auxi = "S" Then
            Vence3.SetFocus
                Else
            Vence2.SetFocus
        End If
    End If
End Sub

Private Sub Vence3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence3.Text, Auxi)
        If Auxi = "S" Then
            Vence4.SetFocus
                Else
            Vence3.SetFocus
        End If
    End If
End Sub

Private Sub Vence4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Vence4.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Vence4.SetFocus
        End If
    End If
End Sub

Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Vence1.Text = "  /  /    "
    Vence2.Text = "  /  /    "
    Vence3.Text = "  /  /    "
    Vence4.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Pesos.Value = True
    Dolares.Value = False
    CtaCte.Value = True
    Documentos.Value = False
    Total.Value = False
    Venci1.Value = True
    Venci2.Value = False
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_CASH
End Sub

