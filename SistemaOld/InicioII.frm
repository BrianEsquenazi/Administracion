VERSION 5.00
Begin VB.Form PrgInicioII 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Stock de Materias Primas"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
End
Attribute VB_Name = "PrgInicioII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call Proceso_Click
End Sub


Private Sub Proceso_Click()

    WDesdeEmpresa = 3
    WHastaEmpresa = 9
    
    PrgInicioII.Hide
    Unload Me
    PrgProc1Auto.Show

End Sub
