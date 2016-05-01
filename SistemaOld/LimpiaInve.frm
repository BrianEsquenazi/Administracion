VERSION 5.00
Begin VB.Form PrgLimpiaInve 
   AutoRedraw      =   -1  'True
   Caption         =   "Limpia la Carga de Inventario"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "PrgLimpiaInve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInventario As Recordset
Dim spInventario As String

Private Sub Cancelar_Click()
    PrgLimpiaInve.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()
    T$ = "Eliminacion del la Carga de Inventario"
    m$ = "!!! ATENCION !!!   Se elimira toda la carga de inventario, Desea Continuar con la operacion      "
    Respuesta% = MsgBox(m$, 32 + 4, T$)
    If Respuesta% = 6 Then
        spInventario = "BorrarInventarioTotal"
        Set rstInventario = db.OpenRecordset(spInventario, dbOpenSnapshot, dbSQLPassThrough)
        Call Cancelar_Click
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLimpiaInve.Caption = "Limpia la Carga de Inventario :  " + !Nombre
        End If
    End With
End Sub
