VERSION 5.00
Begin VB.Form frmLoginDesarrollo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de sesión"
   ClientHeight    =   2955
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1745.912
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   3830.899
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox ClaveIngreso 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   960
      Width           =   2445
   End
   Begin VB.TextBox txtOdbc 
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   1080
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.ComboBox cmbEmpresa 
      Height          =   315
      Left            =   1200
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   480
      Width           =   2580
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   390
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   390
      Left            =   2160
      TabIndex        =   6
      Top             =   1800
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   -120
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Empresa:"
      Height          =   270
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Usuario:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Contraseña:"
      Height          =   270
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   -120
      Visible         =   0   'False
      Width           =   1080
   End
End
Attribute VB_Name = "frmLoginDesarrollo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Rs As Recordset
Dim spConsul As String
Dim gAplicacion As String
Dim strConnect As String
Dim rstOperador As Recordset
Dim spOperador As String

Public LoginSucceeded As Boolean

Private Sub cmbEmpresa_Click()
    Rem txtOdbc = "Empresa" + "0" + Right(Str(cmbEmpresa.ListIndex + 1), 1)
    Rem WEmpresa = "000" + Right(Str(cmbEmpresa.ListIndex + 1), 1)
    Select Case cmbEmpresa.ListIndex
        Case 0
            txtOdbc = "Empresa" + "03"
            WEmpresa = "0003"
        Case 1
            txtOdbc = "Empresa" + "05"
            WEmpresa = "0005"
        Case 2
            txtOdbc = "Empresa" + "04"
            WEmpresa = "0004"
        Case Else
            txtOdbc = "Empresa" + "01"
            WEmpresa = "0001"
    End Select
End Sub

Private Sub cmdCancel_Click()
    'establece la variable global a false
    'para indicar un fallo en el inicio de sesión
    LoginSucceeded = False
    Me.Hide
    End
End Sub

Private Sub cmdOK_Click()
    
    Select Case cmbEmpresa.ListIndex
        Case 0
            txtOdbc = "Empresa" + "03"
            WEmpresa = "0003"
        Case 1
            txtOdbc = "Empresa" + "05"
            WEmpresa = "0005"
        Case 2
            txtOdbc = "Empresa" + "04"
            WEmpresa = "0004"
        Case Else
            txtOdbc = "Empresa" + "01"
            WEmpresa = "0001"
    End Select
    
    WClaveOperador = ClaveIngreso.Text
    
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    spOperador = "ConsultaOperadorClave " + "'" + ClaveIngreso.Text + "'"
    Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
    If rstOperador.RecordCount > 0 Then
        WOperador = rstOperador!Operador
        rstOperador.Close
        Unload frmLoginDesarrollo
        Menu.Show
    End If
    
End Sub

Private Sub Form_Load()

    OPEN_FILE_Empresa
    cmbEmpresa.Clear
    
    txtUserName.Text = "Desarrollo"
    txtPassword.Text = "Desarrollo"
    
    Rem With rstEmpresa
    Rem     .Index = "Empresa"
    Rem     .MoveFirst
    Rem     Do
    Rem         If .EOF = False Then
    Rem              cmbEmpresa.AddItem !Nombre
    Rem             .MoveNext
    Rem                 Else
    Rem             Exit Do
    Rem         End If
    Rem     Loop
    Rem End With
    
    cmbEmpresa.AddItem "SURFACTAN QUIMICOS"
    cmbEmpresa.AddItem "SURFACTAN FARMA"
    cmbEmpresa.AddItem "PELLITAL"
    cmbEmpresa.AddItem "SURFACTAN PIGMENTOS"
    
    cmbEmpresa.ListIndex = 0
    ClaveIngreso.Text = WClaveOperador
    
End Sub

