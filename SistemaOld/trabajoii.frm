VERSION 5.00
Begin VB.Form TRabajoii 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "TRabajoii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLibro As Object
Dim ruta As String

Private Sub Command1_Click()

    Set hexcel = CreateObject("Excel.application")
    ruta = "c:\carlos.xls"

    If Len(Dir(ruta)) > 0 Then
        Set objLibro = hexcel.workbooks.Open(ruta)
        hexcel.Visible = True
            Else
        MsgBox "El archivo no existe"
    End If
    
    Rem HExcel.cells(1, 1).Value = Textt2.Text

End Sub

Private Sub Command2_Click()

    Set AppWord = CreateObject("Word.application")
    ruta = "c:\carlos.doc"

    If Len(Dir(ruta)) > 0 Then
        Set objLibro = AppWord.Documents.Open(ruta)
        'Imprimimos en segundo plano
        AppWord.Documents(1).PrintOut Background
        'Comprobamos que Word no sigue imprimiendo
        Do While AppWord.BackgroundPrintingStatus = 1
        Loop
        'Cerramos el documento sin guardar cambios
        AppWord.Documents.Close (wdDotNotSaveChanges)
        AppWord.Quit
        Set AppWord = Nothing
            Else
        MsgBox "El archivo no existe"
    End If

End Sub






Private Sub adada()


Rem rutina 1



'Asignamos el documento
Set AppWord = CreateObject("word.application")
Set DocWord = AppWord.Documents.Open("C:\hola.doc")
'Colocamos el texto en el marcador
DocWord.Bookmarks("NombreCreador").Select
AppWord.Selection.TypeText Text:=Text1.Text
'Imprimimos en segundo plano
AppWord.Documents(1).PrintOut Background
'Comprobamos que Word no sigue imprimiendo
Do While AppWord.BackgroundPrintingStatus = 1
Loop
'Cerramos el documento sin guardar cambios
AppWord.Documents.Close (wdDotNotSaveChanges)
'Liberamos
Set DocWord = Nothing
'Nos cargamos el objeto creado
AppWord.Quit
Set AppWord = Nothing
End Sub

Eso si no debes olvidar incluir las referencias y componentes necesarios



Rem rutina numero 2

Revisa la propiedad ActivePrinter.




End Sub







