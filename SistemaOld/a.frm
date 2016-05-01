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
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Mensaje, Estilo, Título, Ayuda, Ctxt, Respuesta, MiCadena
Mensaje = "¿Desea continuar?"   ' Define el mensaje.
Estilo = vbYesNo + vbCritical + vbDefaultButton2    ' Define los botones.
Título = "Demostración de MsgBox"   ' Define el título.
Ayuda = "DEMO.HLP"  ' Define el archivo de ayuda.
Ctxt = 1000 ' Define el tema
                ' el contexto
                ' Muestra el mensaje.
Respuesta = MsgBox(Mensaje, Estilo, Título, Ayuda, Ctxt)

If Respuesta = vbYes Then   ' El usuario eligió el botón Sí.
    MiCadena = "Sí" ' Ejecuta una acción.
Else    ' El usuario eligió el botón No.
    MiCadena = "No" ' Ejecuta una acción.
End If

End Sub
