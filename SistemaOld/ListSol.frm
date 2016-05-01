VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListSol 
   Caption         =   "Listado de Solicitudes de Compras Realizadas"
   ClientHeight    =   3825
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   3825
   ScaleWidth      =   5655
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4320
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
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
      Left            =   4440
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "listsol.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Solicitudes de Conpras Realizadas"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListSol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()
    
    With rstPedido
        .Index = "Pedido"
        .MoveFirst
        If .NoMatch = False Then
            Do
                .Edit
                !Pedidonro = Str$(Val(!Pedido))
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    WAno = Right$(Desde.text, 4)
    WMes = Mid$(Desde.text, 4, 2)
    WDia = Left$(Desde.text, 2)
    WDesde = WAno + WMes + WDia
    WAno = Right$(Hasta.text, 4)
    WMes = Mid$(Hasta.text, 4, 2)
    WDia = Left$(Hasta.text, 2)
    Whasta = WAno + WMes + WDia
    Listado.GroupSelectionFormula = "{Pedido.OrdFEntrega} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + Whasta + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    Listado.Action = 1
End Sub

Private Sub Cancela_click()
    With rstPedido
        .Close
    End With
    With rstUnidad
        .Close
    End With
    With rstCliente
        .Close
    End With
    With rstVendedor
        .Close
    End With
    DbsVentas.Close
    Desde.SetFocus
    PrgListSol.Hide
    Menu.SetFocus
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Desde.text, Auxi)
        If Auxi = "S" Then
            Hasta.SetFocus
                Else
            Desde.SetFocus
        End If
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Hasta.text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Hasta.SetFocus
        End If
    End If
End Sub
Sub Form_load()
    Desde.text = "00/00/0000"
    Hasta.text = "31/12/1999"
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

