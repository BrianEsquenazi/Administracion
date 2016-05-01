VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "SISTEMA DE EVALUACION DE PROVEEDORES"
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
      Visible         =   0   'False
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
   Begin VB.Menu dsvsvfsfd 
      Caption         =   "Novedades"
      Begin VB.Menu CalificaGrilla 
         Caption         =   "Actualizacion de Evaluacion Semestral de  Proveedores"
      End
      Begin VB.Menu EvaluaTransportista 
         Caption         =   "Evaluacion de Transportistas"
      End
      Begin VB.Menu EvaluaMantenimiento 
         Caption         =   "Evaluacion de Proveedores de Mantenimiento"
      End
      Begin VB.Menu EvaluaCalibraciones 
         Caption         =   "Evaluacion de Proveedores de Calibraciones"
      End
      Begin VB.Menu EvaluaEnsayos 
         Caption         =   "Evaluacion de Proveedores de Ensayos"
      End
      Begin VB.Menu EvaluaOtros 
         Caption         =   "Evaluacion de Otros Proveedores"
      End
      Begin VB.Menu CalificaGrillaEnvase 
         Caption         =   "Actualizacion de Evaluacion Semestral de  Proveedores de Envases"
      End
   End
   Begin VB.Menu ascfsgfsdf 
      Caption         =   "Listados"
      Begin VB.Menu Califica 
         Caption         =   "Consulta de Evaluacion Semestral Actual de Proveedores"
      End
      Begin VB.Menu ListaCalifica 
         Caption         =   "Planilla de Calculo Teorico de Evaluacion Semestral de Proveedores"
      End
      Begin VB.Menu ListaCheckList 
         Caption         =   "Listado de Check List de Informes de Recepcion"
      End
      Begin VB.Menu ListaEvaluaTransportista 
         Caption         =   "Listado de Evaluacion de Transportista"
      End
      Begin VB.Menu ListaEvaluaServicio 
         Caption         =   "Listado de Evaluacion de Servicio"
      End
      Begin VB.Menu ListaproveRubro 
         Caption         =   "Listado de Proveedores por Rubro"
      End
      Begin VB.Menu ListaVtoCamion 
         Caption         =   "Listado de Vencimiento de Camiones"
      End
      Begin VB.Menu ListaVtoChofer 
         Caption         =   "Listado de Vencimiento de Choferes"
      End
      Begin VB.Menu ListaCheckListExpo 
         Caption         =   "Listado de Check List de Hojas de Ruta"
      End
      Begin VB.Menu CalificaEnvase 
         Caption         =   "Consulta de Evaluacion Semestral Actual de Proveedores de Envases"
      End
      Begin VB.Menu ListaCalificaEnvase 
         Caption         =   "Planilla de Calculo Teorico de Evaluacion Semestral de Proveedores de Envases"
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
Dim rstAtributo As Recordset
Dim spAtributo As String
Dim Atri(10, 100) As Integer

Private Sub asa_Click()
    OPEN_FILE_Esta1
    OPEN_FILE_Esta2
    PrgAscii.Show
End Sub

Private Sub agenda_Click()
    PrgAgendaTotal.Show
End Sub

Private Sub AgendaCargaTarea_Click()
    PrgAgendaCargaTarea.Show
End Sub

Private Sub Califica_Click()
    PrgCalifica.Show
End Sub

Private Sub CalificaEnvase_Click()
    PrgCalificaEnvase.Show
End Sub

Private Sub CalificaGrilla_Click()
    If WMateriaOperador = "S" Then
        PrgCalificaGrilla.Show
    End If
End Sub

Private Sub CalificaGrillaEnvase_Click()
    If WMateriaOperador = "S" Then
        PrgCalificaGrillaEnvases.Show
    End If
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

Private Sub Equipo_Click()
    PrgEquipos.Show
End Sub

Private Sub Cronograma_Click()
    PrgPlanificacion.Show
End Sub

Private Sub CronogramaII_Click()
    PrgCronograma.Show
End Sub

Private Sub Cursadas_Click()
    PrgCursadas.Show
End Sub

Private Sub curso_Click()
    PrgCurso.Show
End Sub

Private Sub Camiones_Click()
    If WTransporteOperador = "S" Then
        PrgCamiones.Show
    End If
End Sub

Private Sub Choferes_Click()
    If WTransporteOperador = "S" Then
        PrgChoferes.Show
    End If
End Sub

Private Sub EvaluaCalibraciones_Click()
    If WSectorOperador > 0 Then
        PrgEvaluaCalibraciones.Show
    End If
End Sub

Private Sub EvaluaEnsayos_Click()
    If WSectorOperador > 0 Then
        PrgEvaluaEnsayos.Show
    End If
End Sub

Private Sub EvaluaMantenimiento_Click()
    If WSectorOperador > 0 Then
        PrgEvaluaMantenimiento.Show
    End If
End Sub

Private Sub EvaluaOtros_Click()
    PrgEvaluaOtros.Show
End Sub

Private Sub EvaluaTransportista_Click()
    If WTransporteOperador = "S" Then
        PrgEvaluaTransporte.Show
    End If
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
                Rem Menu.Caption = "Instrucciones de Produccion (Farma) : " + !Nombre
            End If
        End With
    End If
    
    ZZOperadorResponsable = 1

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

Private Sub Lavado_Click()
    PrgLavado.Show
End Sub

Private Sub MaterialAuxiliar_Click()
    PrgMaterialAuxiliar.Show
End Sub

Private Sub Legajo_Click()
    PrgLegajo.Show
End Sub

Private Sub LegajosVersion_Click()
    PrgLegajoVersion.Show
End Sub

Private Sub ListaCursoCurso_Click()
    PrgListaCursoCurso.Show
End Sub

Private Sub ListaCursoLegajo_Click()
    PrgListaCursoLegajo.Show
End Sub

Private Sub ListaCursoLegajoConsol_Click()
    PrgListaCursoLegajoConsol.Show
End Sub

Private Sub listacursos_Click()
    prglistcursos.Show
End Sub

Private Sub ListaCursoSector_Click()
    PrgListaCursoSector.Show
End Sub

Private Sub Listacursototal_Click()
    PrgListaCursoTotal.Show
End Sub

Private Sub ListaEvaluacionCursos_Click()
    PrgListaEvaluacionCursos.Show
End Sub

Private Sub ListaInactivo_Click()
    PrgListaInactivo.Show
End Sub

Private Sub ListaLegajo_Click()
    PrgListaLegajo.Show
End Sub

Private Sub ListaLegajoCusos_Click()
    PrgListaCursoNoAprobado.Show
End Sub

Private Sub ListaLegajosPerfil_Click()
    PrgListaLegajoPerfil.Show
End Sub

Private Sub ListaPlanifica_Click()
    PrgListaPlanificacion.Show
End Sub

Private Sub ListaPlanilla_Click()
    PrgListaPlanilla.Show
End Sub

Private Sub ListaTareas_Click()
    PrgListaTareas.Show
End Sub

Private Sub Reproceso_Click()
    PrgReproceso.Show
End Sub

Private Sub Sector_Click()
    PrgSector.Show
End Sub

Private Sub tarea_Click()
    PrgTarea.Show
End Sub

Private Sub Tareaversion_Click()
    PrgTareaVersion.Show
End Sub

Private Sub Temas_Click()
    PrgTemas.Show
End Sub

Private Sub Grupo_Click()
    PrgGrupo.Show
End Sub

Private Sub ListaCalifica_Click()
    PrgListaCalifica.Show
End Sub

Private Sub ListaCalificaEnvase_Click()
    PrgListaCalificaEnvase.Show
End Sub

Private Sub ListaCheckList_Click()
    PrgListaCheckList.Show
End Sub

Private Sub ListaCheckListExpo_Click()
    PrgListaCheckListExpo.Show
End Sub

Private Sub ListaEvaluaServicio_Click()
    PrgListaEvaluaServicio.Show
End Sub

Private Sub ListaEvaluaTransportista_Click()
    PrgListaEvaluaTransportista.Show
End Sub

Private Sub ListaproveRubro_Click()
    PrgListaProveRubro.Show
End Sub

Private Sub ListaVtoCamion_Click()
    PrgListaVtoCamion.Show
End Sub

Private Sub ListaVtoChofer_Click()
    PrgListaVtoChofer.Show
End Sub
