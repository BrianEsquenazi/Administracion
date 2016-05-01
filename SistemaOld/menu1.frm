VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Cotizaciones"
   ClientHeight    =   7815
   ClientLeft      =   2430
   ClientTop       =   2175
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
      Caption         =   "Maestros"
      Begin VB.Menu Envases 
         Caption         =   "Ingreso de Envases"
      End
      Begin VB.Menu Arti 
         Caption         =   "Ingreso de Materias Primas"
      End
      Begin VB.Menu Terminado 
         Caption         =   "Ingreso de Producto Terminado"
      End
   End
   Begin VB.Menu Nov 
      Caption         =   "Novedades"
      Begin VB.Menu Cotiza 
         Caption         =   "Ingreso de Cotizaciones"
      End
      Begin VB.Menu Orden 
         Caption         =   "Emision de Ordenes de Compora"
      End
      Begin VB.Menu Informe 
         Caption         =   "Ingreso de Informe de Recepcion"
      End
      Begin VB.Menu Laudo 
         Caption         =   "Ingreso de Laudo de Liberacion"
      End
      Begin VB.Menu Hoja 
         Caption         =   "Ingreso de Hoja de Produccion"
      End
      Begin VB.Menu Movvar 
         Caption         =   "Ingreso de Movimientos Varios"
      End
      Begin VB.Menu MovEnv 
         Caption         =   "Ingreso y Egreso de Envases"
      End
      Begin VB.Menu Pedeti 
         Caption         =   "Emision de Etiquetas de Expostacion"
      End
   End
   Begin VB.Menu listados 
      Caption         =   "Listados"
      Begin VB.Menu ListCot 
         Caption         =   "Listado de Cotizaciones"
      End
      Begin VB.Menu ListOrd 
         Caption         =   "Listado de Ordenes de Compra"
      End
      Begin VB.Menu CotPrv 
         Caption         =   "Listado de Cotizaciones por Proveedor"
      End
      Begin VB.Menu CotArt 
         Caption         =   "Listado de Cotizaciones por Articulo"
      End
      Begin VB.Menu Orden1 
         Caption         =   "Listado de O/C Pend. por Proveedor"
      End
      Begin VB.Menu Orden2 
         Caption         =   "Listado de O/C Pend. por Articulos"
      End
      Begin VB.Menu Listmat1 
         Caption         =   "Listado de Materia Prima"
      End
      Begin VB.Menu Listmat2 
         Caption         =   "Listado de Materia Prima ( Stock )"
      End
      Begin VB.Menu Listter 
         Caption         =   "Listado de Producto Terminado (Stock)"
      End
      Begin VB.Menu ListTer1 
         Caption         =   "Listado de Valuacion de Producto Terminado"
      End
      Begin VB.Menu Minimo 
         Caption         =   "Listado de Materia Prima (Minimo)"
      End
      Begin VB.Menu Minter 
         Caption         =   "Listado de Producto Terminado (Minimo)"
      End
      Begin VB.Menu Compo 
         Caption         =   "Listado de Composicion"
      End
      Begin VB.Menu Proy 
         Caption         =   "Listado de Proyeccion de Entradas"
      End
      Begin VB.Menu FichaMp 
         Caption         =   "Listado de Ficha de Stock de M.P."
      End
      Begin VB.Menu FiechaPt 
         Caption         =   "Listado de Ficha de Stock de P.T."
      End
      Begin VB.Menu Movvar1 
         Caption         =   "Listado de Movimientos Varios de Materia Prima"
      End
      Begin VB.Menu Movvar2 
         Caption         =   "Listado de Movimientos Varios de Producto Terminado"
      End
      Begin VB.Menu Listhoja 
         Caption         =   "Listado de Hojas de Produccion"
      End
   End
   Begin VB.Menu n 
      Caption         =   "Listados "
      Begin VB.Menu CompPrv 
         Caption         =   "Listado de Compras por Proveedor"
      End
      Begin VB.Menu CompMat 
         Caption         =   "Listado de Compras porr Materia Prima"
      End
      Begin VB.Menu ListInf 
         Caption         =   "Listado de Informe de Recepcion"
      End
      Begin VB.Menu ListCont 
         Caption         =   "Listado de Control de Ordenes"
      End
      Begin VB.Menu ConsFichaMp 
         Caption         =   "Consulta de Ficha de Stock M.P."
      End
      Begin VB.Menu ConFichaPt 
         Caption         =   "Consulta de Ficha de Stock P.T."
      End
      Begin VB.Menu Ultima 
         Caption         =   "Listado de Ultima Compra de Materia Prima"
      End
      Begin VB.Menu Eti1 
         Caption         =   "Emision de Etiquetas"
      End
      Begin VB.Menu ListEnv1 
         Caption         =   "Listado de Envases por Cliente"
      End
      Begin VB.Menu ListEnv2 
         Caption         =   "Listado de Envases por Envases"
      End
      Begin VB.Menu Verifica 
         Caption         =   "Listado de Verificacion de Correlatividades"
      End
      Begin VB.Menu Listcomp 
         Caption         =   "Listado de Componentes de Formulas"
      End
   End
   Begin VB.Menu procesos 
      Caption         =   "Procesos"
      Begin VB.Menu CierreStk 
         Caption         =   "Cierre del Stock"
      End
      Begin VB.Menu Proc1 
         Caption         =   "Reproceso de Materia Prima"
      End
      Begin VB.Menu Proc2 
         Caption         =   "Reproceso de Producto Terminado"
      End
      Begin VB.Menu Proc9 
         Caption         =   "Minimo = 0"
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
Private Sub Arti_Click()
    PrgArti.Show
End Sub

Private Sub CompMat_Click()
    PrgOrdart.Show
End Sub

Private Sub Compo_Click()
    PrgCompos.Show
End Sub

Private Sub CompPrv_Click()
    PrgOrdprv.Show
End Sub

Private Sub ConFichaPt_Click()
    PrgConsFicTer.Show
End Sub

Private Sub ConsFichaMp_Click()
    PrgConsFicMat.Show
End Sub

Private Sub CotArt_Click()
    PrgCotart.Show
End Sub

Private Sub Cotiza_Click()
    PrgCoti.Show
End Sub

Private Sub CotPrv_Click()
    PrgCoTPRV.Show
End Sub

Private Sub Envases_Click()
    PrgEnv.Show
End Sub

Private Sub Eti1_Click()
    PrgEti3.Show
End Sub

Private Sub FichaMp_Click()
    PrgFicmat.Show
End Sub

Private Sub FiechaPt_Click()
    PrgFicter.Show
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

Private Sub Listter_Click()
    PrgListter.Show
End Sub

Private Sub ListTer1_Click()
    PrgLister1.Show
End Sub

Private Sub Minimo_Click()
    PrgMinimo.Show
End Sub

Private Sub Minter_Click()
    PrgMinTer.Show
End Sub

Private Sub MovEnv_Click()
    PrgMovEnv.Show
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

Private Sub Proc1_Click()
    PrgProc1.Show
End Sub

Private Sub Proc2_Click()
    PrgProc2.Show
End Sub

Private Sub Proc9_Click()
    PrgProc9.Show
End Sub

Private Sub Proy_Click()
    PrgProyec.Show
End Sub

Private Sub Terminado_Click()
    PrgTermi.Show
End Sub


Private Sub Cambio_Click()
    frmLogin.Show
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
                Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With
            Else
        OPEN_FILE_Empresa
        With rstEmpresa
            .Index = "Empresa"
            .Seek "=", Val(WEmpresa)
            If .NoMatch = False Then
                Menu.Caption = "Sistema de Cotizaciones : " + !Nombre
            End If
        End With
    End If

End Sub

Private Sub Ultima_Click()
    PrgUltima.Show
End Sub

Private Sub Verifica_Click()
    PrgVerifica.Show
End Sub
