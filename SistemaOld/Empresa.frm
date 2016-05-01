VERSION 5.00
Begin VB.Form Empresa 
   AutoRedraw      =   -1  'True
   Caption         =   "Seleccion de Empresas"
   ClientHeight    =   4080
   ClientLeft      =   3165
   ClientTop       =   2655
   ClientWidth     =   6195
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   6195
   Begin VB.CommandButton Command1 
      Caption         =   "Acepta  Empresa"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin VB.ComboBox Selecciona 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   840
      TabIndex        =   0
      Text            =   " "
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Empresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZEmpre As String

Private Sub Command1_Click()
                
    XEmpresa = Selecciona.ListIndex + 1
    WEmpresa = Selecciona.ListIndex + 1
    
    With rstEmpresa
        .Close
    End With
    
    DbsEmpresa.Close
    
    ZZEmpre = Str$(XEmpresa)
    Call Ceros(ZZEmpre, 2)

    txtOdbc = "Empresa" + ZZEmpre
    WEmpresa = "00" + ZZEmpre
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    Empresa.Hide
    Menu.SetFocus

End Sub

Private Sub Form_Load()
    
    OPEN_FILE_Empresa
    Selecciona.Clear

    With rstEmpresa
        .Index = "Empresa"
        .MoveFirst
        Do
           If .EOF = False Then
       
                Selecciona.AddItem !Nombre
                .MoveNext
              Else
               Exit Do
          End If
       
        Loop
    End With
    
End Sub


