VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "SISTEMA DE CAPACITACION"
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
      Begin VB.Menu Sector 
         Caption         =   "Ingreso de Sectores"
      End
      Begin VB.Menu curso 
         Caption         =   "Ingreso de Temas"
      End
      Begin VB.Menu Temas 
         Caption         =   "Ingreso de Cursos"
      End
      Begin VB.Menu Tarea 
         Caption         =   "Ingreso de Perfiles"
      End
      Begin VB.Menu Legajo 
         Caption         =   "Ingreso de Legajos"
      End
      Begin VB.Menu LegajosVersion 
         Caption         =   "Consulta de Version de Legajos"
      End
      Begin VB.Menu Tareaversion 
         Caption         =   "Consulta de Version de Perfiles"
      End
   End
   Begin VB.Menu adsad 
      Caption         =   "Novedades"
      Begin VB.Menu Cronograma 
         Caption         =   "Ingreso de Planificacion Anual de Capacitacion por Legajo"
      End
      Begin VB.Menu CronogramaII 
         Caption         =   "Ingreso de Cronograma de Capacitacion"
      End
      Begin VB.Menu Cursadas 
         Caption         =   "Ingreso de Cursos Realizados"
      End
   End
   Begin VB.Menu sdfdsfd 
      Caption         =   "Listados"
      Begin VB.Menu ListaTareas 
         Caption         =   "Perfil de Puesto"
      End
      Begin VB.Menu ListaLegajo 
         Caption         =   "Informe de Competencia y Necesidades de Capacitacion"
      End
      Begin VB.Menu ListaCursoLegajo 
         Caption         =   "Listado de Temas Realizados por Legajos"
      End
      Begin VB.Menu ListaCursoCurso 
         Caption         =   "Listado de Cursos Realizados por Tema"
      End
      Begin VB.Menu ListaEvaluacionCursos 
         Caption         =   "Listado de Evolucion de Temas Programados"
      End
      Begin VB.Menu ListaLegajoCusos 
         Caption         =   "Listado de Legajos con Necesidades Pendientes por IC y NC vigente"
      End
      Begin VB.Menu ListaCursoSector 
         Caption         =   "Listado de Temas Realizados por Sector"
      End
      Begin VB.Menu ListaPlanifica 
         Caption         =   "Plan de Capacitacion Anual"
      End
      Begin VB.Menu ListaLegajosPerfil 
         Caption         =   "Listado de Legajos por Perfil"
      End
      Begin VB.Menu ListaPlanilla 
         Caption         =   "Planilla de Temas no Programados"
      End
      Begin VB.Menu Listacursototal 
         Caption         =   "Listado de Temas Realizados y No Realizados"
      End
      Begin VB.Menu listacursos 
         Caption         =   "Listado de Temas"
      End
      Begin VB.Menu ListaCursoLegajoConsol 
         Caption         =   "Listado de Temas Realizados por Legajos (Consolidado)"
      End
      Begin VB.Menu ListaInactivo 
         Caption         =   "Listado de Horas Cursadas por Legajo"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu Reproceso 
         Caption         =   "Reproceso de Grabacion  de Cursos"
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
Dim rstAtributo As Recordset
Dim spAtributo As String
Dim Atri(10, 100) As Integer

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
    
    XOperador = Str$(WOperador)
    XProceso = "3"
    WAtributo1 = "00000000000000000000000000000"
    WAtributo2 = "00000000000000000000000000000"
    WAtributo3 = "00000000000000000000000000000"
    WAtributo4 = "00000000000000000000000000000"
    WAtributo5 = "00000000000000000000000000000"
    WAtributo6 = "00000000000000000000000000000"
    WAtributo7 = "00000000000000000000000000000"
    WAtributo8 = "00000000000000000000000000000"
    WAtributo9 = "00000000000000000000000000000"
    WAtributo10 = "00000000000000000000000000000"
    
    XParam = "'" + XOperador + "','" _
                 + XProceso + "'"
    spAtributo = "ConsultaAtributo " + XParam
    Set rstAtributo = db.OpenRecordset(spAtributo, dbOpenSnapshot, dbSQLPassThrough)
    If rstAtributo.RecordCount > 0 Then
        WAtributo1 = rstAtributo!Atributo1 + "00000000000000000000000000000"
        WAtributo2 = rstAtributo!Atributo2 + "00000000000000000000000000000"
        WAtributo3 = rstAtributo!Atributo3 + "00000000000000000000000000000"
        WAtributo4 = rstAtributo!Atributo4 + "00000000000000000000000000000"
        WAtributo5 = rstAtributo!Atributo5 + "00000000000000000000000000000"
        WAtributo6 = rstAtributo!Atributo6 + "00000000000000000000000000000"
        WAtributo7 = rstAtributo!Atributo7 + "00000000000000000000000000000"
        WAtributo8 = rstAtributo!Atributo8 + "00000000000000000000000000000"
        WAtributo9 = rstAtributo!Atributo9 + "00000000000000000000000000000"
        WAtributo10 = rstAtributo!Atributo10 + "00000000000000000000000000000"
        rstAtributo.Close
    End If
    
    For Ciclo = 1 To 10
        Select Case Ciclo
            Case 1
                Auxiliar = WAtributo1
            Case 2
                Auxiliar = WAtributo2
            Case 3
                Auxiliar = WAtributo3
            Case 4
                Auxiliar = WAtributo4
            Case 5
                Auxiliar = WAtributo5
            Case 6
                Auxiliar = WAtributo6
            Case 7
                Auxiliar = WAtributo7
            Case 8
                Auxiliar = WAtributo8
            Case 9
                Auxiliar = WAtributo9
            Case 10
                Auxiliar = WAtributo10
            Case Else
        End Select
        For Ciclo1 = 1 To 30
            aa = Ciclo
            aa1 = Ciclo1
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
        Next Ciclo1
    Next Ciclo
            
    Menu.Sector.Enabled = Atri(1, 1)
    Menu.Curso.Enabled = Atri(1, 2)
    Menu.Tarea.Enabled = Atri(1, 3)
    Menu.Legajo.Enabled = Atri(1, 4)
    Menu.LegajosVersion.Enabled = Atri(1, 5)
    
    Menu.Cronograma.Enabled = Atri(2, 1)
    Menu.CronogramaII.Enabled = Atri(2, 2)
    Menu.Cursadas.Enabled = Atri(2, 3)
    
    Menu.ListaTareas.Enabled = Atri(3, 1)
    Menu.ListaLegajo.Enabled = Atri(3, 2)
    Menu.ListaCursoLegajo.Enabled = Atri(3, 3)
    Menu.ListaCursoCurso.Enabled = Atri(3, 4)
    Menu.ListaEvaluacionCursos.Enabled = Atri(3, 5)
    Menu.ListaLegajoCusos.Enabled = Atri(3, 6)
    
    Menu.Fin.Enabled = 1

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
