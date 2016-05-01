VERSION 5.00
Begin VB.Form Menu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Configuracion de Atributos de los Sistemas"
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
   Begin VB.Menu listados 
      Caption         =   "Opciones"
      Begin VB.Menu ConfigCoti 
         Caption         =   "Configuracion de Cotiza"
      End
      Begin VB.Menu ConfiVentas 
         Caption         =   "Configuracion de Ventas"
      End
      Begin VB.Menu Conficapacitacion 
         Caption         =   "Configuracion de Capacitacion"
      End
      Begin VB.Menu ConfigDesarrollo 
         Caption         =   "Configuracion de Desarrollo"
      End
      Begin VB.Menu configinve 
         Caption         =   "Configuracion de Inversion"
      End
      Begin VB.Menu ConfigLabora 
         Caption         =   "Configuracion de Labora"
      End
      Begin VB.Menu ConfigVende 
         Caption         =   "Configuracion de Vende"
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
Private Sub asa_Click()
    OPEN_FILE_Esta1
    OPEN_FILE_Esta2
    PrgAscii.Show
End Sub

Private Sub Cambio_Click()
    frmLogin.Show
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

Private Sub Conficapacitacion_Click()
    PrgConfigCapacitacion.Show
End Sub

Private Sub ConfigCoti_Click()
    PrgConfigCoti.Show
End Sub

Private Sub ConfigDesarrollo_Click()
    PrgConfigDesarrollo.Show
End Sub

Private Sub configinve_Click()
    PrgConfigInversion.Show
End Sub

Private Sub ConfigLabora_Click()
    PrgConfigLabora.Show
End Sub

Private Sub ConfigVende_Click()
    PrgConfigVende.Show
End Sub

Private Sub ConfiVentas_Click()
    PrgConfigVentas.Show
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
