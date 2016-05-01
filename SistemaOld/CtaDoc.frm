VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCtaCte 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Cuenta Corriente de Clientes"
   ClientHeight    =   6525
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   9015
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   9015
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   3735
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   6495
      Begin VB.Frame Frame4 
         Caption         =   "Moneda"
         Height          =   855
         Left            =   3480
         TabIndex        =   17
         Top             =   1440
         Width           =   2775
         Begin VB.OptionButton Dolares 
            Caption         =   "Dolares"
            Height          =   375
            Left            =   1560
            TabIndex        =   21
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton Pesos 
            Caption         =   "Pesos"
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Comprobantes"
         Height          =   855
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   4815
         Begin VB.OptionButton Total 
            Caption         =   "Total"
            Height          =   255
            Left            =   3000
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Documentos 
            Caption         =   "Documentos"
            Height          =   495
            Left            =   1560
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton CtaCte 
            Caption         =   "Cta. Cte."
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de Listado"
         Height          =   855
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   3135
         Begin VB.OptionButton Tipo2 
            Caption         =   "Completo"
            Height          =   255
            Left            =   1560
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Tipo1 
            Caption         =   "Pendiente"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
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
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   960
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
      Left            =   7440
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ctacte.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Clientes"
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
      Left            =   7800
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   2205
      ItemData        =   "CtaDoc.frx":0000
      Left            =   120
      List            =   "CtaDoc.frx":0007
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   8535
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   7680
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7680
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WPasa As String
Private WTitulo As String
Private Importe3 As Double

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)

    With rstCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                .Edit
                !Importe1 = 0
                !Importe2 = 0
                !Importe3 = 0
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
                        If Pesos.Value = True Then
                            If !Total > 0 Then
                                !Importe1 = !Total
                                !Importe2 = 0
                                    Else
                                !Importe1 = 0
                                !Importe2 = !Total
                            End If
                            !Importe3 = !Saldo
                                Else
                            If !TotalUs > 0 Then
                                !Importe1 = !TotalUs
                                !Importe2 = 0
                                    Else
                                !Importe1 = 0
                                !Importe2 = !TotalUs
                            End If
                            !Importe3 = !SaldoUs
                        End If
                    End If
                End If
                
                Importe3 = !Importe3
                Call Redondeo(Importe3)
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = ""
    
    If CtaCte.Value = True Then
        WTitulo = "Cuenta Corriente - "
    End If
    If Documentos.Value = True Then
        WTitulo = "Documentos - "
    End If
    If Total.Value = True Then
        WTitulo = "Total - "
    End If
    
    If Pesos.Value = True Then
        WTitulo = WTitulo + "Pesos"
    End If
    If Dolares.Value = True Then
        WTitulo = WTitulo + "Dolares"
    End If
    
    With rstAuxiliar
        .Index = "Clave"
        .Seek "=", 1
        If .NoMatch = False Then
            .Edit
            !Nombre = WAuxiliar
            !Auxi1 = ""
            !Varios = Left$(WTitulo, 50)
            .Update
        End If
    End With

    Listado.DataFiles(0) = WEmpresa + "vent.mdb"
    
    Listado.WindowTitle = "Listado de Cuenta Corriente"
    Listado.WindowTop = -3
    Listado.WindowLeft = -3
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Tipo1.Value = True Then
        Listado.GroupSelectionFormula = "{CtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {CtaCte.Importe3} <> 0.00"
            Else
        Listado.GroupSelectionFormula = "{CtaCte.Cliente} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34) + " and {CtaCte.Importe3} <> 999999.99"
    End If
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    With rstClientes
        .Close
    End With
    With rstCtaCte
        .Close
    End With
    DbsVentas.Close
    Desde.SetFocus
    PrgCtaCte.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    With rstClientes
                .Index = "Cliente"
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = !Cliente + "      " + !Razon
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Cliente
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
    End With
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    With rstClientes

                Indice = Pantalla.ListIndex
                Claveven$ = WIndice.List(Indice)
                Desde.Text = Claveven$
                .Index = "Cliente"
                Claveven$ = Desde.Text
                .Seek "=", Claveven$
                If .NoMatch = False Then
                    Desde.Text = !Cliente
                    Hasta.Text = !Cliente
                        Else
                    Desde.Text = Claveven$
                    Hasta.Text = Claveven$
                End If
    End With
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
        Desde.SetFocus
    End If
End Sub
Sub Form_Load()
    Desde.Text = ""
    Hasta.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Tipo1.Value = True
    Tipo2.Value = False
    Pesos.Value = True
    Dolares.Value = False
    CtaCte.Value = True
    Documentos.Value = False
    Total.Value = False
    Frame2.Visible = True
End Sub

