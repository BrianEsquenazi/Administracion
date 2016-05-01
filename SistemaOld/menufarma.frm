VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Instrucciones de Produccion (Farma)"
   ClientHeight    =   7890
   ClientLeft      =   840
   ClientTop       =   795
   ClientWidth     =   10440
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   10440
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu sdf 
      Caption         =   "Maestros"
      Begin VB.Menu Equipo 
         Caption         =   "Ingreso de Equipos Usados"
      End
      Begin VB.Menu MaterialAuxiliar 
         Caption         =   "Ingreso de Materiales Auxiliares"
      End
      Begin VB.Menu Lavado 
         Caption         =   "Ingreso de Metodos de Lavado"
      End
      Begin VB.Menu TextroFijo 
         Caption         =   "Ingreso de Texto Fijo para Procesos"
      End
      Begin VB.Menu CargaI 
         Caption         =   "Ingreso de Equipos a Utilizar en P.T."
      End
      Begin VB.Menu CargaII 
         Caption         =   "Ingreso de Materiales Auxiliares a Utilizar en P.T."
      End
      Begin VB.Menu CargaIII 
         Caption         =   "Ingreso de Instrucciones de Produccion de P.T."
      End
      Begin VB.Menu CargaIIIVersion 
         Caption         =   "Ingreso de Instrucciones de Produccion de P.T. (Version)"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu ImpreCargaI 
         Caption         =   "Impresion del Registro de Produccion de P.T."
      End
      Begin VB.Menu ListaProcesosFarma 
         Caption         =   "Listado de Procesos"
      End
      Begin VB.Menu ImpreCargaIVersion 
         Caption         =   "Impresion del Registro de Produccion de P.T. (Version)"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
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
    frmLoginFarma.Show
End Sub

Private Sub dada_Click()
    prgdada.Show
End Sub

Private Sub esatanu_Click()
    PrgEstaAnu.Show
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

Private Sub Estaven_Click()
    PrgEstaVen.Show
End Sub

Private Sub CargaI_Click()
    PrgCargaI.Show
End Sub

Private Sub Cargaii_Click()
    PrgCargaII.Show
End Sub

Private Sub CargaIII_Click()
    PrgCargaIIIProduccion.Show
End Sub

Private Sub CargaIIIVersion_Click()
    PrgCargaIIIProduccionVersion.Show
End Sub

Private Sub Equipo_Click()
    PrgEquipos.Show
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
                Menu.Caption = "Instrucciones de Produccion (Farma) : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub Listfac_Click()
    PrgListfac.Show
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

Private Sub ImpreCargaI_Click()
    PrgImpreCargaI.Show
End Sub

Private Sub ImpreCargaIVersion_Click()
    PrgImpreCargaIVersion.Show
End Sub

Private Sub Lavado_Click()
    PrgLavado.Show
End Sub

Private Sub ListaProcesosFarma_Click()
    PrgListaProcesosFarma.Show
End Sub

Private Sub MaterialAuxiliar_Click()
    PrgMaterialAuxiliar.Show
End Sub

Private Sub TextroFijo_Click()
    PrgTextoFijo.Show
End Sub
