VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   7890
   ClientLeft      =   630
   ClientTop       =   855
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
      Caption         =   "Menu General"
      Begin VB.Menu Altainv 
         Caption         =   "Ingreso de Talones de Inventario"
      End
      Begin VB.Menu Verital 
         Caption         =   "Listado de Verificacion de Correlatividades de Talones de Inventario"
      End
      Begin VB.Menu verido 
         Caption         =   "Listado de Verificacion de Correlatividades de Talones de Inventario (Duplicados)"
      End
      Begin VB.Menu MovInvMat 
         Caption         =   "Listado de Recuento de Inventario de Matria Prima"
      End
      Begin VB.Menu MovInvTer 
         Caption         =   "Listado de Recuento de Inventario de Producto Terminado"
      End
      Begin VB.Menu DifeInvMat 
         Caption         =   "Listado de Diferencia de Inventario Materia Prima"
      End
      Begin VB.Menu DifeInvTer 
         Caption         =   "Listado de Diferencia de Inventario Producto Terminado"
      End
      Begin VB.Menu MoviMP0 
         Caption         =   "Listado de Inventario con Lote = 0 (MP)"
      End
      Begin VB.Menu MoviPT0 
         Caption         =   "Listado de Inventario con Lote = 0 (PT)"
      End
      Begin VB.Menu DifeInvMatII 
         Caption         =   "Listado de Diferencia de Inventario Materia Prima (Stock Anterior) "
      End
      Begin VB.Menu DifeInvTerII 
         Caption         =   "Listado de Diferencia de Inventario Producto Terminado (Stock Anterior)"
      End
   End
   Begin VB.Menu SDFLÑCVSD 
      Caption         =   "Procesos"
      Begin VB.Menu CierrePrueba 
         Caption         =   "Cierre Prueba"
         Visible         =   0   'False
      End
      Begin VB.Menu LimpiaInve 
         Caption         =   "Limpia Carga de Inventario "
      End
      Begin VB.Menu fdg 
         Caption         =   "Controla Marca de Lote/Existencia Lote"
      End
      Begin VB.Menu CierreStkAnt 
         Caption         =   "Cierre de Stock "
      End
      Begin VB.Menu Cierre 
         Caption         =   "Actualizacion del Stock"
      End
      Begin VB.Menu CierreParcial 
         Caption         =   "Actualizacion Parcial de Pelltal"
         Visible         =   0   'False
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
    PrgAscii.Show
End Sub

Private Sub ActInvMat_Click()

End Sub

Private Sub Altainv_Click()
    PrgAltainv.Show
End Sub

Private Sub Cambio_Click()
    
  Menu.LimpiaInve.Enabled = True
Menu.Cierre.Enabled = True
Menu.CierreStkAnt.Enabled = True
  
    
    frmLogin.Show
End Sub

Private Sub dada_Click()
    prgdada.Show
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

Private Sub Estaven_Click()
    PrgEstaVen.Show
End Sub


Private Sub Cierre_Click()
    PrgCierre.Show
End Sub

Private Sub CierreParcial_Click()
    PrgCierreParcial.Show
End Sub

Private Sub CierrePrueba_Click()
    PrgCierrePrueba.Show
End Sub

Private Sub CierreStkAnt_Click()
    PrgCierreStkAnt.Show
End Sub

Private Sub DifeInvMat_Click()
    OPEN_FILE_Inve
    PrgDifeInvMat.Show
End Sub

Private Sub DifeInvMatII_Click()
    PrgDifeInvMatII.Show
End Sub

Private Sub DifeInvTer_Click()
    OPEN_FILE_Inve
    PrgDifeInvTer.Show
End Sub

Private Sub DifeInvTerII_Click()
    PrgDifeInvTerII.Show
End Sub

Private Sub fdg_Click()
    PrgMiraLote.Show
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


    If UCase(Ingreso) <> "OLULA" And UCase(Ingreso) <> "POLOK" And UCase(Ingreso) <> "GRANADA" Then
        Rem Menu.MovInvTer.Enabled = False
         Menu.LimpiaInve.Enabled = False
         Menu.Cierre.Enabled = False
         Menu.CierreStkAnt.Enabled = False
    End If
Rem DifeInvMatII.Enabled = False
Rem DifeInvTerII.Enabled = False
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

Private Sub Limpia_Click()
    PrgLimpia.Show
End Sub

Private Sub LimpiaInve_Click()
    PrgLimpiaInve.Show
End Sub

Private Sub MoviMP0_Click()
    PrgMoviMp0.Show
End Sub

Private Sub MovInvMat_Click()
    PrgMovInvMat.Show
End Sub

Private Sub MovInvTer_Click()
    PrgMovInvTer.Show
End Sub

Private Sub MoviPT0_Click()
    PrgMoviPt0.Show
End Sub

Private Sub verido_Click()
    PrgVeritalDoble.Show
End Sub

Private Sub Verital_Click()
    PrgVerital.Show
End Sub
