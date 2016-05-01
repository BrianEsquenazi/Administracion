VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prglistcursos 
   Caption         =   "LISTADO DE CURSOS "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4785
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.OptionButton impresora 
      Caption         =   "impresora"
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.OptionButton panta 
      Caption         =   "pantalla"
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.CommandButton aceptar 
      Caption         =   "aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\\193.168.0.2\g$\vb\wlistadocursos.rpt"
      WindowTop       =   600
      WindowHeight    =   600
      WindowBorderStyle=   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "prglistcursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aceptar_Click()

 
 If Impresora.Value = True Then
       Listado.Destination = 1
          Else
         Listado.Destination = 0
    End If
 prglistcursos.Visible = True
  Listado.Action = 1
prglistcursos.Visible = False



End Sub

Private Sub Form_Load()
Rem  Listado.Connect = Connect()
  Rem  Listado.ReportFileName = "Wlistadocursos.rpt"
 Rem Listado.WindowTitle = "Listado de Legajos por Perfil"
 Rem   Listado.WindowTop = 0
 Rem   Listado.WindowLeft = 0
 Rem   Listado.WindowWidth = Screen.Width
Rem Listado.WindowHeight = Screen.Heigh
 
Rem  Listado.WindowState = crptMaximized
Rem prglistcursos.Visible = False
 
Rem Listado.Action = 1

End Sub



