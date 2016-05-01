VERSION 5.00
Begin VB.Form PrgEspe 
   Caption         =   "Especificaciones"
   ClientHeight    =   6795
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4440
      TabIndex        =   26
      Top             =   6140
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Renovar"
      Height          =   300
      Left            =   3360
      TabIndex        =   25
      Top             =   6140
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Actuali&zar"
      Height          =   300
      Left            =   2280
      TabIndex        =   24
      Top             =   6140
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   1200
      TabIndex        =   23
      Top             =   6140
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   120
      TabIndex        =   22
      Top             =   6140
      Width           =   975
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\DevStudio\VB\laboratorio.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Especificaciones"
      Top             =   6450
      Width           =   5520
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor10"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   18
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   21
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo10"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   17
      Left            =   240
      MaxLength       =   50
      TabIndex        =   20
      Top             =   4440
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "valor9"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   16
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo9"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   15
      Left            =   240
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor8"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   14
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   17
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo8"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   240
      MaxLength       =   50
      TabIndex        =   16
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor7"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   15
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo7"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   240
      MaxLength       =   50
      TabIndex        =   14
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   13
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo6"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   240
      MaxLength       =   50
      TabIndex        =   12
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo5"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   240
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   9
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo3"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   240
      MaxLength       =   4
      TabIndex        =   8
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo2"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   240
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Valor1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   3960
      MaxLength       =   50
      TabIndex        =   5
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ensayo1"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   240
      MaxLength       =   4
      TabIndex        =   3
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Producto"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   1
      Top             =   40
      Width           =   3375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Valor"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ensayo"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   380
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Producto:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEspe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  'esto puede generar un error si elimina el último
  'registro o el único registro del recordset
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  'esto sólo es necesario para una aplicación de múltiples usuarios
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Aquí es dónde se sitúa el código de tratamiento de errores
  'Si desea ignorar los errores, comente la línea siguiente
  'Si desea tratarlos, agregue código para controlarlos
  MsgBox "El evento de error de datos generó el error:" & Error$(DataErr)
  Response = 0  'desprecia el error
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  'Esto mostrará la posición actual del registro
  'para dynasets y snapshots
  Data1.Caption = "Registro: " & (Data1.Recordset.AbsolutePosition + 1)
  'para el objeto table debe establecer la propiedad index cuando
  'se cree el recordset y usar la siguiente línea
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'Aquí es donde se pondría el código de validación
  'Se llama a este evento cuando ocurre la siguiente acción
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If KeyAscii = 13 Then
        Rem Data1.Recordset.FIND
        Text2.SetFocus
    End If
End Sub
