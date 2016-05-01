VERSION 5.00
Begin VB.Form Deshabilitar 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeshabilitar 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton cmdHabilitar 
      Caption         =   "Command1"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Deshabilitar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba para deshabilitar el botón cerrar de un formulario         (20/Jun/01)
'
' Basado en un artículo de la KB para VB 3.0 (Q82876)
' How to Disable Close Command in VB Control Menu (System Menu)
'
' ©Guillermo 'guille' Som, 2001
'------------------------------------------------------------------------------
Option Explicit

' Para deshabilitar el menú cerrar (controlbox) de un formulario
Private Declare Function GetSystemMenu Lib "user32" _
    (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long, ByVal wIDNewItem As Long, _
    ByVal lpString As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" _
    (ByVal hWnd As Long) As Long
'
Private Const MF_BYCOMMAND = &H0&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
'
Private Const SC_CLOSE = &HF060&

Private Sub Command1_Click()

End Sub



Private Sub Form_Load()
    ' Deshabilitar el botón de cerrar el formulario
    Dim hMenu As Long
    '
    hMenu = GetSystemMenu(hWnd, 0)
    ' Deshabilitar el menú cerrar del formulario
    Call ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED, -10, "Close")
    '
End Sub

Private Sub cmdDeshabilitar_Click()
    ' Deshabilitar el botón de cerrar el formulario
    Dim hMenu As Long
    '
    hMenu = GetSystemMenu(hWnd, 0)
    ' Deshabilitar el menú cerrar del formulario
    Call ModifyMenu(hMenu, SC_CLOSE, MF_BYCOMMAND Or MF_GRAYED, -10, "Close")
    '
    ' Si esta llamada se hace dentro del Form_Load,
    ' no es necesario redibujar los menús
    ' Redibujar los menús, para que se muestre deshabilitado
    Call DrawMenuBar(hWnd)
    '
End Sub

Private Sub cmdHabilitar_Click()
    ' Habilitar el botón de cerrar el formulario
    Dim hMenu As Long
    '
    hMenu = GetSystemMenu(hWnd, 0)
    ' Esto lo habilita, pero sigue en gris...
    Call ModifyMenu(hMenu, -10, MF_BYCOMMAND Or MF_ENABLED, SC_CLOSE, "Close")
    ' Redibujar los menús, para que se muestre habilitado
    Call DrawMenuBar(hWnd)
    '
End Sub


