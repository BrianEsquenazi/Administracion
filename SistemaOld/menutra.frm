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
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Menu listados 
      Caption         =   "Menu General"
      Begin VB.Menu Traspel 
         Caption         =   "Traspaso de Informacion (Pellital)"
      End
      Begin VB.Menu recpel 
         Caption         =   "Recepcion de Informacion (Pellital)"
      End
      Begin VB.Menu Trassurf 
         Caption         =   "Traspaso de Informacion (Surfactan)"
      End
      Begin VB.Menu recsurf 
         Caption         =   "Recepcion de Informacion (Surfactan)"
      End
      Begin VB.Menu paso3 
         Caption         =   "Traspaso de Ordenes de Compra"
      End
      Begin VB.Menu Paso4 
         Caption         =   "Recepcion de Ordenes de Compra"
      End
      Begin VB.Menu Paso1 
         Caption         =   "Traspaso de datos de ventas"
      End
      Begin VB.Menu Paso2 
         Caption         =   "Recepcion de datos de ventas"
      End
      Begin VB.Menu Prueba 
         Caption         =   "Pruebas Varias"
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

Private Sub Paso1_Click()
    OPEN_FILE_Ctacte
    OPEN_FILE_WCtacte
    PrgPaso1.Show
End Sub

Private Sub Paso2_Click()
    OPEN_FILE_Ctacte
    If Val(WEmpresa) = 2 Then
        OPEN_FILE_WCtacte4
            Else
        OPEN_FILE_WCtacte2
    End If
    PrgPaso2.Show
End Sub

Private Sub Paso3_Click()
    OPEN_FILE_Orden
    OPEN_FILE_WOrden
    PrgPaso3.Show
End Sub

Private Sub Paso4_Click()
    OPEN_FILE_Orden
    OPEN_FILE_WOrden
    PrgPaso4.Show
End Sub

Private Sub Prueba_Click()
    prevar.Show
End Sub

Private Sub recsurf_Click()
    PrgRecsurf.Show

End Sub

Private Sub traspel_Click()
    Prgtraspel.Show
End Sub

Private Sub Cambio_Click()
    Empresa.Show
End Sub


Private Sub Trassurf_Click()
    Prgtrassurf.Show
End Sub
