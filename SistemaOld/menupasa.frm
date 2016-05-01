VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Sistema de Traspaso de Datos"
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
      Begin VB.Menu lee1 
         Caption         =   "Proceso de laboratorio"
      End
      Begin VB.Menu Lee2 
         Caption         =   "Proceso de Cotizaciones"
      End
      Begin VB.Menu lee3 
         Caption         =   "Proceso de Administracion"
      End
      Begin VB.Menu lee4 
         Caption         =   "Proceso de ventas"
      End
      Begin VB.Menu lee5 
         Caption         =   "salva traspaso  mov env// prod ter. // recib hist"
      End
      Begin VB.Menu le6 
         Caption         =   "reproceso de proveedores"
      End
      Begin VB.Menu lee7 
         Caption         =   "Grabacion de mov de envases"
      End
      Begin VB.Menu Lee8 
         Caption         =   "Grabacion de Cotizaciones"
      End
      Begin VB.Menu Lee20 
         Caption         =   "Grabacion de Cotizacviones de Novell"
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

Private Sub Lee20_Click()
    OPEN_FILE_Cotiza
    Prglee20.Show
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

Private Sub lee7_Click()
    OPEN_FILE_MovEnv
    Prglee7.Show
End Sub

Private Sub Lee8_Click()
    OPEN_FILE_Cotiza1
    OPEN_FILE_Cotiza
    Prglee8.Show
End Sub
