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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
    Open "lpt1" For Output As #1

    Print #1, Chr$(27) + Chr$(40) + "19U";
    Print #1, Chr$(27) + Chr$(38) + Chr$(108) + "4" + "H" + "3" + "A";
    Print #1, Chr$(27) + Chr$(40) + Chr$(115) + "10" + Chr$(72)
        
    Print #1, Tab(1); "ESTE IMPORTE ESTA EXPRESADO EN DOLARES ESTADOUNIDENSES."
    Print #1, Tab(1); "REEXPRESION EN PESOS AL SOLO EFECTO CONTABLE/IMPOSITIVO"
    Print #1, Tab(1); "TIPO DE CAMBIO:";
    Print #1, " I.V.A.:";
    Print #1, " TOTAL:";
    Print #1, Tab(1); "IMPORTE NETO:";
    Print #1, Tab(1); "CONDICIONES : SI POR FUERZA MAYOR NO FUESE POSIBLE EL"
    Print #1, Tab(1); "PAGO EN DOLARES BILLETE; SURFACTAN S.A. PODRA OPTAR EN"
    Print #1, Tab(1); "RECIBIR PESOS GLOBAL/08 COTIZACION MERCADO NVA.YORK, EN"
    Print #1, Tab(1); "CANTIDAD SUFICIENTE PARA  ADQUIRIR EL  EQUIVALENTE AL"
    Print #1, Tab(1); "PRECIO EN DOLARES. SI EL IMPORTE NO SE CANCELARA EN EL"
    Print #1, Tab(1); "PLAZO ESTIPULADO A PARTIR DE SU VENCIMIENTO Y HASTA LA"
    Print #1, Tab(1); "FECHA EFVO.PAGO SE APLICARA UNA TASA DEL 20.00% ANUAL"
    Print #1, ""
    Print #1, ""
    Print #1, ""


    
    Print #1, Chr$(12)

    Close #1
    
    
    Close
    End

End Sub
