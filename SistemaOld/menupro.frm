VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Proceso de Stock"
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
   Begin VB.Menu listados 
      Caption         =   "Menu General"
      Begin VB.Menu Proc1 
         Caption         =   "Reproceso de Materia Prima"
      End
      Begin VB.Menu Proc2 
         Caption         =   "Reproceso de Producto Terminado"
      End
      Begin VB.Menu Proc3 
         Caption         =   "Cierre de Stock de Materia Prima y Producto Terminado"
      End
      Begin VB.Menu Proc4 
         Caption         =   "Blanque de campos"
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
Private Sub Fin_Click()
    Close
    End
    
End Sub

Private Sub le6_Click()
    OPEN_FILE_Proveedor
    Prglee6.Show
End Sub

Private Sub lee1_Click()
    OPEN_FILE_ENSAYOS
    OPEN_FILE_ESPECIFICACIONES
    OPEN_FILE_ESPECIF
    OPEN_FILE_PRUEBA
    OPEN_FILE_PrueTer
    Prglee1.Show
End Sub

Private Sub traspa_Click()
    OPEN_FILE_ENSAYOS
    OPEN_FILE_ESPECIFICACIONES
    OPEN_FILE_ESPECIF
    OPEN_FILE_PRUEBA
    OPEN_FILE_PrueTer
    OPEN_FILE_Movlab
    OPEN_FILE_Cotiza
    OPEN_FILE_Hoja
    OPEN_FILE_Informe
    OPEN_FILE_LAUDO
    OPEN_FILE_Orden
    OPEN_FILE_Movvar
    OPEN_FILE_Articulo
    OPEN_FILE_Clientes
    OPEN_FILE_Composicion
    OPEN_FILE_Ctacte
    OPEN_FILE_DescComp
    OPEN_FILE_Precios
    Prgtraspa.Show
End Sub

Private Sub Form_Activate()

    If WEmpresa = "" Then
        WEmpresa = "0005"
    End If

End Sub

Private Sub lee2_Click()
    OPEN_FILE_Hoja
    OPEN_FILE_Cotiza
    OPEN_FILE_Informe
    OPEN_FILE_Orden
    OPEN_FILE_LAUDO
    OPEN_FILE_Movvar
    Prglee2.Show
End Sub

Private Sub lee3_Click()
    OPEN_FILE_Banco
    OPEN_FILE_Cuenta
    OPEN_FILE_Proveedor
    OPEN_FILE_CtaCtePrv
    OPEN_FILE_Ivacomp
    OPEN_FILE_Imputac
    OPEN_FILE_Depositos
    OPEN_FILE_Recibos
    OPEN_FILE_Pagos
   Prglee3.Show
End Sub

Private Sub lee4_Click()
    OPEN_FILE_LINEAS
    OPEN_FILE_Rubros
    OPEN_FILE_Vendedores
    OPEN_FILE_ENVASES
    OPEN_FILE_Clientes
    OPEN_FILE_TERMINADO
    OPEN_FILE_Precios
    OPEN_FILE_Articulo
    OPEN_FILE_Composicion
    OPEN_FILE_Estadistica
    OPEN_FILE_Ctacte
    OPEN_FILE_Pago
    OPEN_FILE_Pedido
    Prglee4.Show
End Sub

Private Sub lee5_Click()
    OPEN_FILE_MovEnv
    OPEN_FILE_TERMINADO
    OPEN_FILE_Recibos
    Prglee5.Show
End Sub

Private Sub Proc1_Click()
    OPEN_FILE_Orden
    OPEN_FILE_Proveedor
    OPEN_FILE_Articulo
    OPEN_FILE_LAUDO
    OPEN_FILE_Hoja
    OPEN_FILE_Movvar
    OPEN_FILE_Movlab
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    PrgProc1.Show
End Sub

Private Sub Proc2_Click()
    OPEN_FILE_TERMINADO
    OPEN_FILE_Hoja
    OPEN_FILE_Movvar
    OPEN_FILE_Movlab
    OPEN_FILE_Estadistica
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    PrgProc2.Show
End Sub

Private Sub Proc3_Click()
    OPEN_FILE_LAUDO
    OPEN_FILE_Hoja
    OPEN_FILE_Movvar
    OPEN_FILE_Movlab
    OPEN_FILE_Estadistica
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    PrgProc3.Show
End Sub

Private Sub Proc4_Click()
    OPEN_FILE_Hoja
    OPEN_FILE_Movvar
    OPEN_FILE_Movlab
    OPEN_FILE_Estadistica
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
    PrgProc4.Show
End Sub
