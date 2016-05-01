VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Tablero de Comando"
   ClientHeight    =   7815
   ClientLeft      =   2280
   ClientTop       =   780
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   7350
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Cambio 
      Caption         =   "Cambio de Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Menu Maestros 
      Caption         =   "Menu General"
      Begin VB.Menu ProyAtraso 
         Caption         =   "Proyeccion de Atrasos"
         Visible         =   0   'False
      End
      Begin VB.Menu calculapt 
         Caption         =   "Grabacion de Stock y Costos de Productos Terminados"
      End
      Begin VB.Menu CalculaMp 
         Caption         =   "Grabacion de Stock y Costos de Dy  / DW"
      End
      Begin VB.Menu ListaComando 
         Caption         =   "Listado de Comando"
      End
      Begin VB.Menu FinCot 
         Caption         =   "Fin del Sistema"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Actualiza_Click()
    PrgModped.Show
End Sub

Private Sub Arti_Click()
    PrgArti.Show
End Sub

Private Sub CierreStk_Click()
    OPEN_FILE_InveMp
    OPEN_FILE_InvePt
    PrgCierre.Show
End Sub

Private Sub CompMat_Click()
    PrgOrdart.Show
End Sub

Private Sub Compo_Click()
    PrgCompos.Show
End Sub

Private Sub Compos1_Click()
    PrgCompos1.Show
End Sub

Private Sub CompPrv_Click()
    PrgOrdprv.Show
End Sub

Private Sub ConArtCon_Click()
    PrgOrdartCon.Show
End Sub

Private Sub ConFichaPt_Click()
    PrgConsFicTer.Show
End Sub

Private Sub ConsFichaMp_Click()
    PrgConsFicMat.Show
End Sub

Private Sub ConsFicMatAnt_Click()
    PrgConsFicMatAnt.Show
End Sub

Private Sub ConsFicTerAnt_Click()
    PrgConsFicTerAnt.Show
End Sub

Private Sub ConsumoArt_Click()
    PrgConsumoArt.Show
End Sub

Private Sub Consumoter_Click()
    PrgConsumoTer.Show
End Sub

Private Sub Costo_Click()
    PrgCosto.Show
End Sub

Private Sub Cotart_Click()
    PrgCotart.Show
End Sub

Private Sub Cotiza_Click()
    PrgCoti.Show
End Sub

Private Sub Cotprv_Click()
    PrgCoTPRV.Show
End Sub

Private Sub Entdev_Click()
    PrgEntdev.Show
End Sub

Private Sub Envases_Click()
    PrgEnv.Show
End Sub

Private Sub Eti1_Click()
    OPEN_FILE_Etiqueta
    OPEN_FILE_Empresa
    PrgEti3.Show
End Sub

Private Sub FichaMp_Click()
    PrgFicmat.Show
End Sub

Private Sub FiechaPt_Click()
    PrgFicter.Show
End Sub

Private Sub CalculaMp_Click()
    PrgCalculaMp.Show
End Sub

Private Sub calculapt_Click()
    PrgCalculaPt.Show
End Sub

Private Sub FinCot_Click()
    Close
    End
End Sub

Private Sub Hoja_Click()
    PrgHoja.Show
End Sub

Private Sub Informe_Click()
    PrgInforme.Show
End Sub

Private Sub Laudo_Click()
    Prglaudo.Show
End Sub

Private Sub Listcomp_Click()
    PrgListcomp.Show
End Sub

Private Sub Listcomp1_Click()
    PrgListcomp1.Show
End Sub

Private Sub ListCont_Click()
    PrgControl.Show
End Sub

Private Sub ListCot_Click()
    PrgListcot.Show
End Sub

Private Sub ListEnv1_Click()
    PrgListmov1.Show
End Sub

Private Sub ListEnv2_Click()
    PrgListmov2.Show
End Sub

Private Sub Listhoja_Click()
    PrgListhoja.Show
End Sub

Private Sub ListInf_Click()
    PrgListinf.Show
End Sub

Private Sub Listmat1_Click()
    PrgListmat1.Show
End Sub

Private Sub Listmat2_Click()
    PrgStkmat.Show
End Sub

Private Sub ListOrd_Click()
    PrgListOrd.Show
End Sub

Private Sub Listpres_Click()
    PrgListpres.Show
End Sub

Private Sub Listter_Click()
    PrgListter.Show
End Sub

Private Sub ListTer1_Click()
    PrgLister1.Show
End Sub

Private Sub Lotemat_Click()
    PrgLotemat.Show
End Sub

Private Sub Loteter_Click()
    PrgLoteter.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub Minimo1_Click()
    PrgMinimoConsol.Show
End Sub

Private Sub Minter_Click()
    PrgMinTer.Show
End Sub

Private Sub MInter1_Click()
    PrgMinTerConsol.Show
End Sub

Private Sub mirasol_Click()
    PrgMIrasol.Show
End Sub

Private Sub MovEnv_Click()
    PrgMovEnv.Show
End Sub

Private Sub Movgas_Click()
    PrgMovgas.Show
End Sub

Private Sub Movguia_Click()
    PrgMovguia.Show
End Sub

Private Sub Movvar_Click()
    PrgMovvar.Show
End Sub

Private Sub Movvar1_Click()
    PrgMovvar1.Show
End Sub

Private Sub Movvar2_Click()
    PrgMovvar2.Show
End Sub

Private Sub Orden_Click()
    PrgOrden.Show
End Sub

Private Sub Orden1_Click()
    PrgOrdPenPrv.Show
End Sub

Private Sub Orden2_Click()
    PrgOrdPenArt.Show
End Sub

Private Sub Pedeti_Click()
    PrgPedeti.Show
End Sub

Private Sub Pedpen_Click()
    PrgPedPen.Show
End Sub

Private Sub Prestamo_Click()
    PrgPrestamo.Show
End Sub

Private Sub Proc1_Click()
    PrgProc1.Show
End Sub

Private Sub Proc101_Click()
    PrgProc101.Show
End Sub

Private Sub Proc102_Click()
    PrgProc102.Show
End Sub

Private Sub Proc11_Click()
    PrgProc11.Show
End Sub

Private Sub Proc2_Click()
    PrgProc2.Show
End Sub

Private Sub Proc9_Click()
    PrgProc9.Show
End Sub

Private Sub ProcHoja_Click()
    PrgProchoja.Show
End Sub

Private Sub prove_Click()
    PrgProve.Show
End Sub

Private Sub Proy_Click()
    PrgProyec.Show
End Sub

Private Sub Sedronar_Click()
    OPEN_FILE_Sedro
    PrgSedronar.Show
End Sub

Private Sub Solic_Click()
    PrgSolic.Show
End Sub

Private Sub Terminado_Click()
    PrgTermi.Show
End Sub

Private Sub Cambio_Click()
    frmLogin1.Show
End Sub

Private Sub Fin_Click()
    Menu.WindowState = 1
End Sub

Private Sub Form_Activate()

    If WEmpresa = "" Then
        WEmpresa = "0001"
        Rem Empresa.Show
        Rem Empresa.SetFocus
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Rem Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Rem Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub Ultima_Click()
    PrgUltima.Show
End Sub

Private Sub valo1_Click()
    PrgStock1.Show
End Sub

Private Sub Valo2_Click()
    PrgStock2.Show
End Sub

Private Sub Verifica_Click()
    PrgVerifica.Show
End Sub

Private Sub verilot1_Click()
    PrgVerilot1.Show
End Sub

Private Sub Verilot2_Click()
    PrgVerilot2.Show
End Sub

Private Sub verilot3_Click()
    PrgVerilot3.Show
End Sub

Private Sub ListaComando_Click()
    If Val(WEmpresa) = 1 Then
        PrgListaComando.Show
            Else
        PrgListaComandoPelli.Show
    End If
End Sub
