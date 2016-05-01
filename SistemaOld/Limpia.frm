VERSION 5.00
Begin VB.Form PrgLimpia 
   AutoRedraw      =   -1  'True
   Caption         =   "Cierre de Stock"
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
Attribute VB_Name = "PrgLimpia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInventario As Recordset
Dim spInventario As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Private Uno As String
Private Dos As String
Private Tres As String
Private Auxi As String
Private Auxi1 As String
Private Auxi2 As String
Private WArticulo As String
Private WTerminado As String
Private WLote As String
Private WCantidad As String
Private WLiberada As String
Private WMarca As String
Private WSaldo As String
Dim Vector(5000, 10) As String

Private Sub Cancelar_Click()

    PrgLimpia.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    WFecha = "01/01/1990"

    spArticulo = "ModificaArticuloInicial0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    spTerminado = "ModificaTerminadoInicial0"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)

    
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgLimpia.Caption = "limpia :  " + !Nombre
        End If
    End With

End Sub
