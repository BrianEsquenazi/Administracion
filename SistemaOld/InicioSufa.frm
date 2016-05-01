VERSION 5.00
Begin VB.Form PrgInicioSurfa 
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
Attribute VB_Name = "PrgInicioSurfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Call Proceso_Click
End Sub


Private Sub Proceso_Click()

    WDesdeEmpresa = 1
    WHastaEmpresa = 11
    
    WDesdeEmpresa = 1
    WHastaEmpresa = 1
    
    PrgInicioSurfa.Hide
    Unload Me
    Rem PrgProc1Auto.Show
    PrgVerilot1AUTO.Show
    
End Sub
