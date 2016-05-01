VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Instrucciones de Procesos de Fabricacion"
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
      Begin VB.Menu Operarios 
         Caption         =   "Ingreso de Operarios"
      End
      Begin VB.Menu cargacontrol 
         Caption         =   "Orden de Fabricacion"
      End
      Begin VB.Menu HojaNueva 
         Caption         =   "Carga  de Hoja de Produccion"
         Visible         =   0   'False
      End
      Begin VB.Menu ConsultaHojaII 
         Caption         =   "Asignacion de Hoja de Produccion"
         Visible         =   0   'False
      End
      Begin VB.Menu ConsuktaHoja 
         Caption         =   "Procesamiento de Hojas de Produccion"
         Visible         =   0   'False
      End
      Begin VB.Menu ConsukltaHojaEnvasa 
         Caption         =   "Verificacion de Hojas para Envasamiento"
         Visible         =   0   'False
      End
      Begin VB.Menu CONSULTAEQUI 
         Caption         =   "Supervision de Produccion"
         Visible         =   0   'False
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
Dim WFecha As Date

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

Private Sub Command1_Click()


    OPEN_FILE_Temperatura0

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura0
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                aa1 = !Hora
                aa2 = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    Stop


End Sub

Private Sub cargacontrol_Click()
    PrgCargaControl.Show
End Sub

Private Sub ConsukltaHojaEnvasa_Click()
    PrgConsultaHojaEnvasado.Show
End Sub

Private Sub ConsuktaHoja_Click()
    PrgConsultaHoja.Show
End Sub

Private Sub CONSULTAEQUI_Click()
    PrgPanelControl.Show
End Sub

Private Sub ConsultaHojaII_Click()
    PrgConsultaHojaII.Show
End Sub

Private Sub EquipoFabrica_Click()
    PrgEquiposFabrica.Show
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
