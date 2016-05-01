VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Prgverif 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Iva Compras"
   ClientHeight    =   3825
   ClientLeft      =   3240
   ClientTop       =   2025
   ClientWidth     =   5655
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.Frame Frame2 
      Caption         =   "Control Listado"
      Height          =   1455
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   3735
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
         Left            =   1680
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4440
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ivacomp.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva Compras"
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
      Left            =   4320
      TabIndex        =   6
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   1425
      ItemData        =   "verif.frx":0000
      Left            =   480
      List            =   "verif.frx":0007
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   1560
      TabIndex        =   4
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "Prgverif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    With rstCtaCtePrv
    
        .Index = "CtaCte"
        .MoveFirst
        
        Do
            
            If Val(!Tipo) = 1 Then
            
            WProveedor = !Proveedor
            WLetra = !Letra
            WTipo = !Tipo
            WPunto = !Punto
            WNumero = !Numero
            WFecha = !Fecha
            
            If Val(Left$(WFechaord, 6)) >= 199910 Then
            
            With rstCtaCtePrv
                .Index = "CtaCte"
                .Seek "=", WProveedor + WLetra + WTipo + WPunto + WNumero
                If .NoMatch Then
                    .AddNew
                    !Proveedor = WProveedor
                    !Letra = WLetra
                    !Tipo = WTipo
                    !Punto = WPunto
                    !Numero = WNumero
                    !Fecha = WFecha
                    !Estado = "1"
                    !Vencimiento = Wvencimiento
                    !Vencimiento1 = WVencimiento1
                    !NroInterno = wnrointerno
                    
                    !Total = WTotal
                    !Saldo = 0
                    
                    !Clave = WProveedor + WLetra + WTipo + WPunto + WNumero
                    !OrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                    !OrdVencimiento = Right$(Wvencimiento, 4) + Mid$(Wvencimiento, 4, 2) + Left$(Wvencimiento, 2)
                    Select Case Val(WTipo)
                        Case 1
                            !Impre = "FC"
                        Case 2
                            !Impre = "ND"
                        Case 3
                            !Impre = "NC"
                            !Total = !Total * -1
                            !Saldo = !Saldo * -1
                        Case Else
                            !Impre = ""
                    End Select
                    !Empresa = 1
                    .Update
                End If
            End With
            
            End If
            
            End If
            
            .MoveNext
            If .EOF = True Then
                Exit Do
            End If
            
        Loop
            
    End With
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstCtaCtePrv
        .Close
    End With
    With rstIvaComp
        .Close
    End With
    DbsAdminis.Close
    Desde.SetFocus
    Prgverif.Hide
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
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

