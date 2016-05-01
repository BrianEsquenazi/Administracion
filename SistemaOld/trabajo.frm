VERSION 5.00
Begin VB.Form TRabajo 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   3120
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   975
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "TRabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLibro As Object
Dim Ruta As String

Private Sub Command1_Click()

WOrigen = "C:\vb5.exe"
WDEstino = "A:\vb5.exe"

FileCopy WOrigen, WDEstino

End Sub

Private Sub Command2_Click()

    Set appWord = CreateObject("Word.application")
    Ruta = "c:\hoja.doc"

    If Len(Dir(Ruta)) > 0 Then
    
        Set objLibro = appWord.documents.Open(Ruta)
        
        appWord.Selection.Font.Size = 18
        appWord.Selection.TypeText Chr$(9) + Chr$(9) + Chr$(9) + Chr$(9) + Chr$(9) + "Hoja de Seguridad" + Chr$(13)
        appWord.Selection.Font.Size = 11
        appWord.Selection.TypeText "Fecha : 25.05.04" + Chr$(13)
        appWord.Selection.TypeText "Nombre Comercial : FACTRNA MDF" + Chr$(13)
        
        Rem ENTER APPWord.Selection.TypeParagraph
        
        appWord.Visible = True
Stop
        Rem appWord.documents(1).Printout Background
        Rem 'Comprobamos que Word no sigue imprimiendo
        Rem Do While appWord.BackgroundPrintingStatus = 1
        Rem Loop
        
        appWord.documents.Close (wdDotNotSaveChanges)
        appWord.Quit
        Set appWord = Nothing
        
        Close
        End
        
            Else
        MsgBox "El archivo no existe"
    End If

End Sub

Private Sub Command82_Click()

    Set appWord = CreateObject("Word.application")
    Ruta = "c:\carlos.doc"

    If Len(Dir(Ruta)) > 0 Then
        Set objLibro = appWord.documents.Open(Ruta)
        w1.Selection.TypeText txtTypeText.Text
        'Imprimimos en segundo plano
        appWord.documents(1).Printout Background
        'Comprobamos que Word no sigue imprimiendo
        Do While appWord.BackgroundPrintingStatus = 1
        Loop
        'Cerramos el documento sin guardar cambios
        appWord.documents.Close (wdDotNotSaveChanges)
        appWord.Quit
        Set appWord = Nothing
            Else
        MsgBox "El archivo no existe"
    End If

End Sub




Private Sub Command3_Click()

    Call ShellAbout(Me.hWnd, "david esquenazi", "Copyright 1997, Empresa res system ", Me.Icon)

End Sub

Private Sub Command4_Click()
        
        Stop
        
        With appWord
            .Selection.Find.ClearFormatting
            .Selection.Find.Replacement.ClearFormatting
            With .Selection.Find
                Rem .Text = ChrW(8220) & "nombre 1" & ChrW(8221)
                .Text = "DADA1"
                .Replacement.Text = "Prueba"
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            .Selection.Find.Execute Replace:=wdReplaceAll
        End With

End Sub



Private Sub Command5_Click()
Dim appWord As Object
'crea objeto word
Set appWord = CreateObject("Word.Application")
appWord.Visible = True
appWord.documents.Open "c:\tem\prueba.txt"
appWord.Selection.WholeStory
appWord.Selection.Font.Size = 8
appWord.ActiveDocument.PageSetup.Orientation = 1
appWord.ActiveDocument.Printout
'appWord.Quit True
'Set objWord = Nothing
End Sub

Private Sub Command6_Click()
objWord.SaveAs "MiArchivo.doc"
End Sub


