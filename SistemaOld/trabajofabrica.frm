VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Canti8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Canti8.Text = Pusing("###,###.##", Canti8.Text)
        Call Canti8_DblClick
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Canti8_DblClick()
    
    If tipo8.Text = "M" Then
        ZTipo = tipo8.Text
        ZArticulo = Mp8.Text
        ZCantidad = Canti8.Text
        If ZArticulo <> "  -   -   " And Val(ZCantidad) <> 0 Then
            ZLugar = 8
            Call Inicio_Carga
        End If
            Else
        If tipo8.Text = "T" Then
            ZTipo = tipo8.Text
            ZTerminado = Pt8.Text
            ZCantidad = Canti8.Text
            If ZTerminado <> "  -   -   " And Val(ZCantidad) <> 0 Then
                ZLugar = 8
                Call Inicio_Carga
            End If
        End If
    End If

End Sub

