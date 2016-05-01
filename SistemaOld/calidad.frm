VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Menu sfdfdg 
      Caption         =   "Sistemna de Calidad"
      Begin VB.Menu prv 
         Caption         =   "Proveedores"
      End
      Begin VB.Menu fdghfg 
         Caption         =   "Clientes"
         Begin VB.Menu dfg 
            Caption         =   "Reclamos"
         End
      End
      Begin VB.Menu cvbvc 
         Caption         =   "Productos"
         Begin VB.Menu xzvcxcv 
            Caption         =   "Codigos"
         End
         Begin VB.Menu wqrew 
            Caption         =   "Metodos"
         End
         Begin VB.Menu sdfsdf 
            Caption         =   "Procesos"
         End
      End
      Begin VB.Menu wetrew 
         Caption         =   "Documentos"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub prv_Click()
    List1.Visible = True
    List1.Clear
    List1.AddItem "Calificacion"
    List1.AddItem "Evaluacion"
    List1.AddItem "Resultado"
End Sub

