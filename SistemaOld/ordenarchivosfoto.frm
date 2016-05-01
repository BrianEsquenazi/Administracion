VERSION 5.00
Begin VB.Form PrgOrdenArchivosFoto 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Archivos Adjuntados a la OC"
   ClientHeight    =   8610
   ClientLeft      =   3465
   ClientTop       =   525
   ClientWidth     =   8790
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8610
   ScaleWidth      =   8790
   Visible         =   0   'False
   Begin VB.Image Foto 
      Height          =   8415
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   8535
   End
End
Attribute VB_Name = "PrgOrdenArchivosFoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ass = WWWPasaFoto
    Foto.Picture = LoadPicture(WWWPasaFoto)
End Sub

