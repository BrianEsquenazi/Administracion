VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Ventas"
   ClientHeight    =   6375
   ClientLeft      =   2430
   ClientTop       =   2175
   ClientWidth     =   7350
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6375
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
      Begin VB.Menu repro 
         Caption         =   "precios por  lcliente"
      End
      Begin VB.Menu fgcbhg 
         Caption         =   "estadistica"
      End
      Begin VB.Menu Rec2 
         Caption         =   "Reproceso de Recibos"
      End
      Begin VB.Menu pag2 
         Caption         =   "Reproceso de Op:pagos"
      End
      Begin VB.Menu verif 
         Caption         =   "Verificacion de facturas en catcte nuervo"
      End
      Begin VB.Menu repto2 
         Caption         =   "verifica duplicaiocn en ctacte"
      End
      Begin VB.Menu veri1 
         Caption         =   "verifica ctacte prov saldos"
      End
      Begin VB.Menu veri2 
         Caption         =   "veri ctacte clientes saldos"
      End
      Begin VB.Menu veriant 
         Caption         =   "Verifica anricipos"
      End
      Begin VB.Menu repro3 
         Caption         =   "Cambia anticipos"
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
Private Sub Arti_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Articulo
    OPEN_FILE_ENVASES
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgArti.Show
End Sub

Private Sub Camb_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Cambios
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCambios.Show
End Sub

Private Sub Cash_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCash.Show
End Sub

Private Sub Clientes_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Vendedores
    OPEN_FILE_Rubros
    OPEN_FILE_Pago
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    prgcliente.Show
End Sub

Private Sub Compo_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    OPEN_FILE_Composicion
    OPEN_FILE_TERMINADO
    OPEN_FILE_Articulo
    Rem rem rem menu.hide
    PrgCompo.Show
End Sub

Private Sub CtaCte1_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    OPEN_FILE_Vendedores
    OPEN_FILE_Rubros
    OPEN_FILE_Pago
    OPEN_FILE_Recibos
    Rem rem rem menu.hide
    PrgCtaCte1.Show
End Sub

Private Sub CtaCteCli_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgCtaCte.Show
End Sub

Private Sub Devol_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Numero
    OPEN_FILE_Cambios
    OPEN_FILE_Precios
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Pedido
    OPEN_FILE_Ctacte
    OPEN_FILE_Estadistica
    OPEN_FILE_Pago
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgDevol.Show
End Sub

Private Sub Envases_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_ENVASES
    OPEN_FILE_Auxiliar
    Rem rem rem menu.hide
    PrgEnv.Show
End Sub

Private Sub Factura_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Numero
    OPEN_FILE_Cambios
    OPEN_FILE_Precios
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Pedido
    OPEN_FILE_ENVASES
    OPEN_FILE_Ctacte
    OPEN_FILE_Estadistica
    OPEN_FILE_Pago
    OPEN_FILE_Auxiliar
    Rem menu.hide
    PrgFactu.Show
End Sub

Private Sub IvaVentas_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgIvaven.Show
End Sub

Private Sub Lineas_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_LINEAS
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgLinea.Show
End Sub



Private Sub Modif_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Precios
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgModif.Show
End Sub

Private Sub Pedido_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Precios
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Pedido
    OPEN_FILE_ENVASES
    OPEN_FILE_Pago
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgPedido.Show
End Sub

Private Sub PedPen_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Pedido
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgPedPen.Show
End Sub

Private Sub Precios_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Precios
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgPrecio.Show
End Sub

Private Sub RECI_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    OPEN_FILE_Vendedores
    OPEN_FILE_Rubros
    OPEN_FILE_Pago
    OPEN_FILE_Recibos
    Rem rem rem menu.hide
    PrgRec.Show

End Sub

Private Sub Rubros_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Rubros
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgRubro.Show
End Sub

Private Sub SalCtaCteCli_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgSaldoCta.Show
End Sub

Private Sub Terminado_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_TERMINADO
    OPEN_FILE_LINEAS
    OPEN_FILE_ENVASES
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgTermi.Show
End Sub

Private Sub Ultima_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Articulo
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgUltima.Show
End Sub

Private Sub Varios_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Numero
    OPEN_FILE_Cambios
    OPEN_FILE_Clientes
    OPEN_FILE_Ctacte
    OPEN_FILE_DescComp
    OPEN_FILE_Pago
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgVarios.Show
End Sub

Private Sub Vendedores_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Vendedores
    OPEN_FILE_Auxiliar
    Rem rem menu.hide
    PrgVendedor.Show
End Sub

Private Sub pago_Click()
    OPEN_FILE_Auxiliar
    OPEN_FILE_Empresa
    OPEN_FILE_Pago
    Rem rem menu.hide
    PrgCondPago.Show
End Sub

Private Sub Cambio_Click()
    Empresa.Show
End Sub

Private Sub fgcbhg_Click()
    OPEN_FILE_Estadistica
    OPEN_FILE_TERMINADO
    Rem rem menu.hide
    Prgrepro1.Show
End Sub

Private Sub Fin_Click()
    Close
    End
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

Private Sub pag2_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Proveedor
    OPEN_FILE_Pagos
    OPEN_FILE_CtaCtePrv
    Rem rem menu.hide
    Prgpag2.Show
End Sub

Private Sub Rec2_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Clientes
    OPEN_FILE_Recibos
    OPEN_FILE_Ctacte
    Rem rem menu.hide
    PrgRec2.Show
End Sub

Private Sub repro_Click()
    OPEN_FILE_Precios
    Rem rem menu.hide
    Prgrepro.Show
End Sub

Private Sub repro3_Click()
    OPEN_FILE_Ctacte
    Rem rem menu.hide
    Prgrepro3.Show
End Sub

Private Sub repto2_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Ctacte
    OPEN_FILE_Orden
    OPEN_FILE_Pagos
    Rem rem menu.hide
    Prgrepro2.Show
End Sub

Private Sub veri1_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Pagos
    OPEN_FILE_CtaCtePrv
    Rem rem menu.hide
    PrgVeri1.Show
End Sub

Private Sub veri2_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Estadistica
    OPEN_FILE_DescComp
    OPEN_FILE_Ctacte
    Rem rem menu.hide
    PrgVeri2.Show
End Sub

Private Sub veriant_Click()
    OPEN_FILE_Recibos
    OPEN_FILE_Ctacte
    Rem rem menu.hide
    Prgveriant.Show
End Sub

Private Sub verif_Click()
    OPEN_FILE_Empresa
    OPEN_FILE_Ivacomp
    OPEN_FILE_CtaCtePrv
    Rem rem menu.hide
    Prgverif.Show
End Sub
