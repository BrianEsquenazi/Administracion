VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Hoja de ruta "
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
   Begin VB.Menu sdf 
      Caption         =   "Maestros"
      Begin VB.Menu Camiones 
         Caption         =   "Ingreso de Camiones"
      End
      Begin VB.Menu Choferes 
         Caption         =   "Ingreso de Choferes"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Novedades"
      Begin VB.Menu HojaRuta 
         Caption         =   "Hoja de Ruta"
      End
      Begin VB.Menu ConfirmaHoja 
         Caption         =   "Confirmacion de la Hoja de Ruta"
      End
      Begin VB.Menu ConsultaHojaRuta 
         Caption         =   "Consulta de Hoja de Ruta (Cot)"
         Visible         =   0   'False
      End
      Begin VB.Menu ConsultaHojaRutaII 
         Caption         =   "Consulta de Hoja de Ruta (Cliente)"
      End
      Begin VB.Menu ListaIncumplimientoHoja 
         Caption         =   "Listados de Incumplimientos de Entrega"
      End
      Begin VB.Menu ListaVtoCamion 
         Caption         =   "Listado de Vencimiento de Camiones"
      End
      Begin VB.Menu ListaVtoChofer 
         Caption         =   "Listado de Vencimiento de Choferes"
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
    frmLoginIII.Show
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
    PrgCargaIII.Show
End Sub

Private Sub Equipo_Click()
    PrgEquipos.Show
End Sub

Private Sub CargaIv_Click()
    PrgCargaIV.Show
End Sub

Private Sub CargaIvVersion_Click()
    PrgCargaIVVersion.Show
End Sub

Private Sub CARGANUEVA_Click()
    PrgCargaNueva.Show
End Sub

Private Sub ConsukltaHojaEnvasa_Click()
    PrgConsultaHojaEnvasado.Show
End Sub

Private Sub ConsuktaHoja_Click()
    PrgConsultaHoja.Show
End Sub

Private Sub CONSULTAEQUI_Click()
    PrgConsultaHojaTotal.Show
End Sub

Private Sub ConsultaHojaII_Click()
    PrgConsultaHojaII.Show
End Sub

Private Sub EquipoFabrica_Click()
    PrgEquiposFabrica.Show
End Sub

Private Sub Camiones_Click()
    PrgCamiones.Show
End Sub

Private Sub Choferes_Click()
    PrgChoferes.Show
End Sub

Private Sub ConfirmaHoja_Click()
    PrgHojaRutaConfirmacion.Show
End Sub

Private Sub ConsultaHojaRuta_Click()
    PrgConsultaHojaRuta.Show
End Sub

Private Sub ConsultaHojaRutaII_Click()
    PrgConsultaHojaRutaII.Show
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
                Menu.Caption = "Instrucciones de Procesos de Fabricacion : " + !Nombre
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

Private Sub HojaNueva_Click()
    PrgHojaNueva.Show
End Sub

Private Sub ImpreCargaI_Click()
    PrgImpreCargaFabrica.Show
End Sub

Private Sub MaterialAuxiliar_Click()
    PrgMaterialAuxiliar.Show
End Sub

Private Sub LeePlanilla_Click()
    PrgLeePlanilla.Show
End Sub

Private Sub ListaProcesos_Click()
    PrgListaProcesos.Show
End Sub

Private Sub MetodoEnvasa_Click()
    PrgMetodoFiltrado.Show
End Sub

Private Sub Operarios_Click()
    PrgOperarios.Show
End Sub

Private Sub HojaRuta_Click()
    PrgHojaRuta.Show
End Sub

Private Sub ListaIncumplimientoHoja_Click()
    PrgListaIncumplimientoHoja.Show
End Sub

Private Sub ListaVtoCamion_Click()
    PrgListaVtoCamion.Show
End Sub

Private Sub ListaVtoChofer_Click()
    PrgListaVtoChofer.Show
End Sub
