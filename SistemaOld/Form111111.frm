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
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1800
      TabIndex        =   1
      Top             =   1920
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem Option Explicit
  
Function Imprimir(Path As String, _
                  Optional Visible_Word As Boolean = True) As Boolean
  
       
    ' variable de objeto para acceder al Word
    Dim Obj_Word As Object
       
           
    ' crea el objeto
    Set Obj_Word = CreateObject("Word.Application")
       
       
    ' Visible / No visible
    If Visible_Word Then
        Obj_Word.Visible = True
    Else
        Obj_Word.Visible = False
    End If
           
    'Abre el documento
    Obj_Word.Documents.Open Path
       
    ' Imprime el documento activo con Printout
    Obj_Word.ActiveDocument.Printout
  
    ' Cierra el documento
    Obj_Word.Quit
       
    ' Elimina la referencia
    Set Obj_Word = Nothing
       
    ' retorno
    If Err.Number = 0 Then
        Imprimir = True
    End If
  
  
Exit Function
  
Error_Function:
  
' error
MsgBox Err.Description
On Error Resume Next
  
Set Obj_Word = Nothing
Obj_Word.Quit
  
End Function
  
' CommandButton que imprime
''''''''''''''''''''''''''''''''''''''''''''''''''
  
Private Sub Command1_Click()
    Dim ret As Boolean
       
    ' le pasa el documento de word que se va a imprimir
    ret = Imprimir("c:\pdfprint\dada.txt", False)
       
    If ret Then
       MsgBox "Ok", vbInformation
    End If
  
End Sub

Private Sub Form_Load()
    Command1.Caption = " Imprimir Documento "
End Sub
  
Private Sub Command2_Click()
    Dim Header, I, Y    ' Declara variables.
    
    Print .PaperBin = 1
    Print "Imprimiendo..."  ' Avisa en el formulario.
    Header = "Demostración de impresión - Página "  ' Establece la cadena de encabezado.
    For I = 1 To 3
        Printer.Print Header;   ' Imprime el encabezado.
        Printer.Print Printer.Page  ' Imprime el número de página.
        Y = Printer.CurrentY + 10   ' Establece la posición de la línea.
        ' Dibuja una línea atravesando la página.
        Printer.Line (0, Y)-(Printer.ScaleWidth, Y)     ' Dibuja una línea.

        For K = 1 To 50
            Printer.Print String(K, " ");   ' Imprime una cadena de espacios.
            Printer.Print "Visual Basic ";  ' Imprime texto.
            Printer.Print Printer.Page  ' Imprime el número de página.
        Next
        Printer.NewPage
    Next I
    Printer.EndDoc
    End




End Sub


