VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgIvacmp 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Iva por Comprobantes"
   ClientHeight    =   4230
   ClientLeft      =   855
   ClientTop       =   2250
   ClientWidth     =   10665
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   10665
   Begin VB.TextBox Impo55 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   43
      Text            =   " "
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Impo54 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      TabIndex        =   42
      Text            =   " "
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Impo53 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   41
      Text            =   " "
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Impo52 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   40
      Text            =   " "
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Impo45 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   39
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Impo35 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   38
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Impo25 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   37
      Text            =   " "
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Impo15 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7080
      TabIndex        =   36
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Impo44 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      TabIndex        =   35
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Impo34 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      TabIndex        =   34
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Impo24 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      TabIndex        =   33
      Text            =   " "
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Impo14 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5760
      TabIndex        =   32
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Impo43 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   31
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Impo33 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   30
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Impo23 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   29
      Text            =   " "
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Impo13 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   28
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Impo42 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   27
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Impo32 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   26
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Impo22 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   25
      Text            =   " "
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Impo12 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   24
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Impo51 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   23
      Text            =   " "
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox Impo41 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      Text            =   " "
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Impo31 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      Text            =   " "
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Impo21 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Text            =   " "
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Impo11 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Text            =   " "
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1320
         TabIndex        =   9
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
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4920
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ivaprv.rpt"
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
      Left            =   4800
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   120
      TabIndex        =   44
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "No Grabado"
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label9 
      Caption         =   "Iva10,5%"
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Iva21%"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Neto"
      Height          =   255
      Left            =   3360
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Total"
      Height          =   255
      Left            =   1920
      TabIndex        =   14
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Exportaciones"
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notas de Debito"
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Notas de Credito"
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label factura 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Facturas"
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "PrgIvacmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WImpo11 As Double
Private WImpo12 As Double
Private WImpo13 As Double
Private WImpo14 As Double
Private WImpo15 As Double
Private WImpo21 As Double
Private WImpo22 As Double
Private WImpo23 As Double
Private WImpo24 As Double
Private WImpo25 As Double
Private WImpo31 As Double
Private WImpo32 As Double
Private WImpo33 As Double
Private WImpo34 As Double
Private WImpo35 As Double
Private WImpo41 As Double
Private WImpo42 As Double
Private WImpo43 As Double
Private WImpo44 As Double
Private WImpo45 As Double
Private WImpo51 As Double
Private WImpo52 As Double
Private WImpo53 As Double
Private WImpo54 As Double
Private WImpo55 As Double


Private Sub Acepta_Click()


    WImpo11 = 0
    WImpo12 = 0
    WImpo13 = 0
    WImpo14 = 0
    WImpo15 = 0
    WImpo21 = 0
    WImpo22 = 0
    WImpo23 = 0
    WImpo24 = 0
    WImpo25 = 0
    WImpo31 = 0
    WImpo32 = 0
    WImpo33 = 0
    WImpo34 = 0
    WImpo35 = 0
    WImpo41 = 0
    WImpo42 = 0
    WImpo43 = 0
    WImpo44 = 0
    WImpo45 = 0
    WImpo51 = 0
    WImpo52 = 0
    WImpo53 = 0
    WImpo54 = 0
    WImpo55 = 0

    WAno = Right$(Desde.Text, 4)
    WMes = Mid$(Desde.Text, 4, 2)
    WDia = Left$(Desde.Text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.Text, 4)
    WMes = Mid$(Hasta.Text, 4, 2)
    WDia = Left$(Hasta.Text, 2)
    WHasta = WAno + WMes + WDia

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    WTitulo = "del " + Desde.Text + " al " + Hasta.Text
    
    With rstCtacte
            .Index = "Clave"
            .MoveFirst
            Do
                If !OrdFecha >= WDesde And !OrdFecha <= WHasta Then
                If !Tipo >= 1 And !Tipo <= 5 Then
                
                    WCliente = !Cliente
                                
                    With rstClientes
                        .Index = "Cliente"
                        Claveven$ = WCliente
                        .Seek "=", WCliente
                        If .NoMatch = False Then
                            WProvincia = !Provincia
                        End If
                    End With
                    
                    If WProvincia <> 23 Then
                        Select Case !Tipo
                            Case 1, 3
                                WImpo13 = WImpo13 + !Iva1
                                WImpo14 = WImpo14 + !Iva2
                                If !Iva1 <> 0 Then
                                    WImpo12 = WImpo12 + !Neto
                                        Else
                                    WImpo15 = WImpo15 + !Neto
                                End If
                                WImpo11 = WImpo11 + !Neto + !Iva1 + !Iva2
                                
                            Case 2, 5
                                WImpo23 = WImpo23 + !Iva1
                                WImpo24 = WImpo24 + !Iva2
                                If !Iva1 <> 0 Then
                                    WImpo22 = WImpo22 + !Neto
                                        Else
                                    WImpo25 = WImpo25 + !Neto
                                End If
                                WImpo21 = WImpo21 + !Neto + !Iva1 + !Iva2
                                
                            Case Else
                                WImpo33 = WImpo33 + !Iva1
                                WImpo34 = WImpo34 + !Iva2
                                If !Iva1 <> 0 Then
                                    WImpo32 = WImpo32 + !Neto
                                        Else
                                    WImpo35 = WImpo35 + !Neto
                                End If
                                WImpo31 = WImpo31 + !Neto + !Iva1 + !Iva2
                        End Select
                    
                            Else

                        WImpo43 = WImpo43 + !Iva1
                        WImpo44 = WImpo44 + !Iva2
                        If !Iva1 <> 0 Then
                            WImpo42 = WImpo42 + !Neto
                                Else
                            WImpo45 = WImpo45 + !Neto
                        End If
                        WImpo41 = WImpo41 + !Neto + !Iva1 + !Iva2
                    End If
                    
                    WImpo53 = WImpo53 + !Iva1
                    WImpo54 = WImpo54 + !Iva2
                    If !Iva1 <> 0 Then
                        WImpo52 = WImpo52 + !Neto
                            Else
                        WImpo55 = WImpo55 + !Neto
                    End If
                    WImpo51 = WImpo51 + !Neto + !Iva1 + !Iva2
                    
                    
                    
                End If
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    Impo11.Text = Alinea("###,###.##", Str$(WImpo11))
    Impo12.Text = Alinea("###,###.##", Str$(WImpo12))
    Impo13.Text = Alinea("###,###.##", Str$(WImpo13))
    Impo14.Text = Alinea("###,###.##", Str$(WImpo14))
    Impo15.Text = Alinea("###,###.##", Str$(WImpo15))
    
    Impo21.Text = Alinea("###,###.##", Str$(WImpo21))
    Impo22.Text = Alinea("###,###.##", Str$(WImpo22))
    Impo23.Text = Alinea("###,###.##", Str$(WImpo23))
    Impo24.Text = Alinea("###,###.##", Str$(WImpo24))
    Impo25.Text = Alinea("###,###.##", Str$(WImpo25))
    
    Impo31.Text = Alinea("###,###.##", Str$(WImpo31))
    Impo32.Text = Alinea("###,###.##", Str$(WImpo32))
    Impo33.Text = Alinea("###,###.##", Str$(WImpo33))
    Impo34.Text = Alinea("###,###.##", Str$(WImpo34))
    Impo35.Text = Alinea("###,###.##", Str$(WImpo35))
    
    Impo41.Text = Alinea("###,###.##", Str$(WImpo41))
    Impo42.Text = Alinea("###,###.##", Str$(WImpo42))
    Impo43.Text = Alinea("###,###.##", Str$(WImpo43))
    Impo44.Text = Alinea("###,###.##", Str$(WImpo44))
    Impo45.Text = Alinea("###,###.##", Str$(WImpo45))
        
    Impo51.Text = Alinea("###,###.##", Str$(WImpo51))
    Impo52.Text = Alinea("###,###.##", Str$(WImpo52))
    Impo53.Text = Alinea("###,###.##", Str$(WImpo53))
    Impo54.Text = Alinea("###,###.##", Str$(WImpo54))
    Impo55.Text = Alinea("###,###.##", Str$(WImpo55))

    If Impresora.Value = True Then
        Open "lpt1" For Output As #1
        Print #1, WAuxiliar
        Print #1, WTitulo
        
        Print #1, ""
        Print #1, Tab(1); "Comprobante";
        Print #1, Tab(20); "Total";
        Print #1, Tab(30); "Neto";
        Print #1, Tab(40); "Iva Inc%";
        Print #1, Tab(50); "Iva ";
        Print #1, Tab(60); "Exento"
        
        Print #1, ""
        Print #1, Tab(1); "Facturas";
        Print #1, Tab(20); Impo11.Text;
        Print #1, Tab(30); Impo12.Text;
        Print #1, Tab(40); Impo13.Text;
        Print #1, Tab(50); Impo14.Text;
        Print #1, Tab(60); Impo15.Text
        
        Print #1, Tab(1); "N.Credito";
        Print #1, Tab(20); Impo21.Text;
        Print #1, Tab(30); Impo22.Text;
        Print #1, Tab(40); Impo23.Text;
        Print #1, Tab(50); Impo24.Text;
        Print #1, Tab(60); Impo25.Text
        
        Print #1, Tab(1); "N.Debito";
        Print #1, Tab(20); Impo31.Text;
        Print #1, Tab(30); Impo32.Text;
        Print #1, Tab(40); Impo33.Text;
        Print #1, Tab(50); Impo34.Text;
        Print #1, Tab(60); Impo35.Text
        
        Print #1, Tab(1); "Export.";
        Print #1, Tab(20); Impo41.Text;
        Print #1, Tab(30); Impo42.Text;
        Print #1, Tab(40); Impo43.Text;
        Print #1, Tab(50); Impo44.Text;
        Print #1, Tab(60); Impo45.Text
        
        Print #1, ""
        Print #1, Tab(1); "Total";
        Print #1, Tab(20); Impo51.Text;
        Print #1, Tab(30); Impo52.Text;
        Print #1, Tab(40); Impo53.Text;
        Print #1, Tab(50); Impo54.Text;
        Print #1, Tab(60); Impo55.Text
    End If
    
End Sub

Private Sub Cancela_click()
    With rstClientes
        .Close
    End With
    With rstCtacte
        .Close
    End With
    DbsVentas.Close
    Desde.SetFocus
    PrgIvacmp.Hide
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
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub
Sub Form_Load()
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Impo11.Text = ""
    Impo12.Text = ""
    Impo13.Text = ""
    Impo14.Text = ""
    Impo15.Text = ""
    Impo21.Text = ""
    Impo22.Text = ""
    Impo23.Text = ""
    Impo24.Text = ""
    Impo25.Text = ""
    Impo31.Text = ""
    Impo32.Text = ""
    Impo33.Text = ""
    Impo34.Text = ""
    Impo35.Text = ""
    Impo41.Text = ""
    Impo42.Text = ""
    Impo43.Text = ""
    Impo44.Text = ""
    Impo45.Text = ""
    Impo51.Text = ""
    Impo52.Text = ""
    Impo53.Text = ""
    Impo54.Text = ""
    Impo55.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

