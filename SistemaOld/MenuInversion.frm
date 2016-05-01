VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Proyectos de Inversion"
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
      Caption         =   "Menu"
      Begin VB.Menu Sector 
         Caption         =   "Ingreso de Sectores"
      End
      Begin VB.Menu Proyecto 
         Caption         =   "Ingreso de Proyectos"
      End
      Begin VB.Menu Asigna 
         Caption         =   "Asignacion de Proyectos al Año"
      End
      Begin VB.Menu avance 
         Caption         =   "Ingreso de Avance de los Proyectos"
      End
      Begin VB.Menu avanceproy 
         Caption         =   "avance de proyectos"
      End
      Begin VB.Menu ListaAvance 
         Caption         =   "Listado de Avance de Proyecto"
      End
      Begin VB.Menu aprobacion 
         Caption         =   "Aprobacin de proyectos"
      End
      Begin VB.Menu listacondicion 
         Caption         =   "Listado estado de proyectos"
      End
      Begin VB.Menu grafic 
         Caption         =   "grafico"
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

Private Sub aprobacion_Click()
Prgaprobacion.Show
End Sub

Private Sub Asigna_Click()
    PrgAsigna.Show
End Sub

Private Sub Avance_Click()
    PrgAvance.Show
End Sub

Private Sub avancecodigo_Click()
proyectocostos.Show
End Sub

Private Sub avanceproy_Click()
  miraproyecto.Show
End Sub

Private Sub Cambio_Click()
    frmLoginInversion.Show
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
        Rem OPEN_FILE_Empresa
        Rem With rstEmpresa
        Rem .Index = "Empresa"
        Rem     .Seek "=", Val(WEmpresa)
        Rem     If .NoMatch = False Then
        Rem         Menu.Caption = "Instrucciones de Procesos de Fabricacion : " + !Nombre
        Rem    End If
        Rem End With
    End If
    
    
    XOperador = Str$(woperador)
    XProceso = "5"
    WAtributo1 = "00000000000000000000000000000000000000000000"
    WAtributo2 = "00000000000000000000000000000000000000000000"
    WAtributo3 = "00000000000000000000000000000000000000000000"
    WAtributo4 = "00000000000000000000000000000000000000000000"
    WAtributo5 = "00000000000000000000000000000000000000000000"
    WAtributo6 = "00000000000000000000000000000000000000000000"
    WAtributo7 = "00000000000000000000000000000000000000000000"
    WAtributo8 = "00000000000000000000000000000000000000000000"
    WAtributo9 = "00000000000000000000000000000000000000000000"
    WAtributo10 = "00000000000000000000000000000000000000000000"
    
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
        For Ciclo1 = 1 To 31
            Atri(Ciclo, Ciclo1) = Val(Mid$(Auxiliar, Ciclo1, 1))
        Next Ciclo1
    Next Ciclo
            
    Menu.Sector.Enabled = Atri(1, 1)
    Menu.Proyecto.Enabled = Atri(1, 2)
    Menu.Asigna.Enabled = Atri(1, 3)
    Menu.avance.Enabled = Atri(1, 4)
    Menu.ListaAvance.Enabled = Atri(1, 5)

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

Private Sub grafic_Click()
GRAFICO.Show
End Sub

Private Sub jurgen_Click()
Form1.Show
End Sub

Private Sub ListaAvance_Click()
    PrgListaInversion.Show
End Sub

Private Sub listacondicion_Click()
prgestado.Show

End Sub

Private Sub Proyecto_Click()
    PrgProyectos.Show
End Sub

Private Sub Sector_Click()
    PrgSectores.Show
End Sub
