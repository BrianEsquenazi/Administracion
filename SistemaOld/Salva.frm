VERSION 5.00
Begin VB.Form PrgSalva 
   AutoRedraw      =   -1  'True
   Caption         =   "Procesos Varios"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11655
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   975
      Left            =   2640
      TabIndex        =   1
      Top             =   3120
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   3135
   End
End
Attribute VB_Name = "PrgSalva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Private WEntradas As Double
Private WSalidas As Double
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim XParam As String

Sub Cancelar_Click()
    PrgSalva.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()
        
    Rem For a = 1441 To 1459
    Rem     spEstadistica = "Borrarestadistica1 " + "'" + Str$(a) + "'"
    Rem     Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    Rem Next a
        
    a = 404143
        
    spMovvar = "BorrarMovvar " + "'" + Str$(a) + "'"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgSalva.Caption = "Procesos VArios :  " + !Nombre
        End If
    End With

End Sub
