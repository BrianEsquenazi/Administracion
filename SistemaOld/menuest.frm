VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Estadistica"
   ClientHeight    =   7890
   ClientLeft      =   840
   ClientTop       =   795
   ClientWidth     =   10440
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   10440
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu Estaven 
         Caption         =   "1.-Estadistica de ventas por Vendedor, Rubo y Linea"
      End
      Begin VB.Menu Esta1 
         Caption         =   "2.-Estadistica de ventas por rubro y cliente"
      End
      Begin VB.Menu Esta2 
         Caption         =   "3.-Estadistica de ventas por linea y producto (Ind.)"
      End
      Begin VB.Menu Esta3 
         Caption         =   "4.-Estadistica de ventas por linea y prodcuto"
      End
      Begin VB.Menu Esta4 
         Caption         =   "5.-Estadistica de ventas por vendedor, cliente y linea"
      End
      Begin VB.Menu Esta5 
         Caption         =   "6.-Estadistica de ventas por cliente"
      End
      Begin VB.Menu Esta6 
         Caption         =   "7.-Estadistica de ventas por vendedor"
      End
      Begin VB.Menu Esta7 
         Caption         =   "8.-Estadisrtica de ventas por producto"
      End
      Begin VB.Menu Rancli 
         Caption         =   "9.-Ranking por Cliente"
      End
      Begin VB.Menu Ranpro 
         Caption         =   "10.-Ranking por producto"
      End
      Begin VB.Menu Ranlin 
         Caption         =   "11.-Ranking por Linea"
      End
      Begin VB.Menu Listfac 
         Caption         =   "12.-Listados de Facturas"
      End
      Begin VB.Menu esatanu 
         Caption         =   "13.-Listado de Estadisticas Anuales"
      End
      Begin VB.Menu esatanudy 
         Caption         =   "14.-Listado de Estadisticas Anuales por Tipo"
      End
      Begin VB.Menu EstaAnuClie 
         Caption         =   "15.-Listado de Estadisticas Anuales por Cliente"
      End
      Begin VB.Menu EstaExpo 
         Caption         =   "16.-Listado de Estadisticas de Exportacion"
      End
      Begin VB.Menu CargaProyeccionventa 
         Caption         =   "17.-Carga de Proyeccion de Venta"
      End
      Begin VB.Menu ProyeccionMp 
         Caption         =   "18.-Listado de Proyeccion de Consumo de M.P."
      End
      Begin VB.Menu EstaanuInter 
         Caption         =   "19 - Listado de Estadisticas InterAnuales"
      End
      Begin VB.Menu Moreno 
         Caption         =   "Moreno"
      End
      Begin VB.Menu listasalva 
         Caption         =   "lsiatsalva"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu asa 
         Caption         =   "Conversion  de Estadisticas a Ascii"
      End
      Begin VB.Menu Fin 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub asa_Click()
    OPEN_FILE_Esta1
    OPEN_FILE_Esta2
    PrgAscii.Show
End Sub

Private Sub Cambio_Click()
    frmLogin.Show
End Sub

Private Sub dada_Click()
    prgdada.Show
End Sub

Private Sub CargaProyeccionventa_Click()
    PrgCargaProyeccionVenTA.Show
End Sub

Private Sub controlventas_Click()
    PrgEstaAnuControl.Show
End Sub

Private Sub esatanu_Click()
    PrgEstaAnu.Show
End Sub

Private Sub esatanudy_Click()
    PrgEstaAnuDy.Show
End Sub

Private Sub Esta1_Click()
    PrgEsta1.Show
End Sub

Private Sub Esta2_Click()
    PrgEsta2.Show
End Sub

Private Sub Esta3_Click()
    PrgEsta3.Show
End Sub

Private Sub Esta4_Click()
    PrgEsta4.Show
End Sub

Private Sub Esta5_Click()
    PrgEsta5.Show
End Sub

Private Sub Esta6_Click()
    PrgEsta6.Show
End Sub

Private Sub Esta7_Click()
    PrgEsta7.Show
End Sub

Private Sub EstaAnuClie_Click()
    PrgEstaAnuClie.Show
End Sub

Private Sub EstaanuInter_Click()
    PrgEstaAnuInter.Show
End Sub

Private Sub EstaExpo_Click()
    PrgEstaExpo.Show
End Sub

Private Sub Estaven_Click()
    PrgEstaVen.Show
End Sub

Private Sub Fin_Click()
    Close
    End
    Rem Menu.WindowState = 1
End Sub

Private Sub Form_Activate()

    If WEmpresa = "" Then
        WEmpresa = "0001"
    End If

    If WEmpresa = "" Then
        Empresa.Show
        Empresa.SetFocus
        WEmpresa = 1
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de ventas : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub listasalva_Click()
    PrgEstasalva.Show
End Sub

Private Sub Listfac_Click()
    PrgListfac.Show
End Sub

Private Sub Moreno_Click()
    PrgMoreno.Show
End Sub

Private Sub ProyeccionMp_Click()
    PrgProyeccionMp.Show
End Sub

Private Sub Rancli_Click()
    PrgRankClie.Show
End Sub

Private Sub Ranpro_Click()
    PrgRankProd.Show
End Sub

Private Sub Ranlin_Click()
    PrgRankLIn.Show
End Sub

Private Sub SalvaPrecios_Click()
    PrgSalvaPrecios.Show
End Sub
