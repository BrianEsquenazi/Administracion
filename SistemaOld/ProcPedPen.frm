VERSION 5.00
Begin VB.Form PrgProcPedPen 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Pedidos Pendientes"
   ClientHeight    =   3165
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   3165
   ScaleWidth      =   8145
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
End
Attribute VB_Name = "PrgProcPedPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Tres As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim WVector(1000, 5) As String
Dim LugarVector As Integer
Dim WTipopro As String

Private Sub Acepta_Click()

    spPedido = "ModificaPedpen0"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    spArticulo = "ModificaArticuloVenta0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    spTerminado = "ModificaTerminadoPedido0"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)

    XParam = "'" + "00000000" + "','" _
                 + "99999999" + "'"
    spPedido = "ModificaPedpen " + XParam
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase WVector
    LugarVector = 0
    
    spPedido = "ListaPedidoPend"
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    EntraVector = "S"
                    For Ciclo = 1 To LugarVector
                        If WVector(Ciclo, 1) = rstPedido!Terminado Then
                            WVector(Ciclo, 2) = Str$(Val(WVector(Ciclo, 2)) + rstPedido!Importe)
                            EntraVector = "N"
                            Exit For
                        End If
                    Next Ciclo
                    If EntraVector = "S" Then
                        LugarVector = LugarVector + 1
                        WVector(LugarVector, 1) = rstPedido!Terminado
                        WVector(Ciclo, 2) = Str$(rstPedido!Importe)
                    End If
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
    For Ciclo = 1 To LugarVector
        WProducto = WVector(Ciclo, 1)
        WTipopro = Left$(WProducto, 2)
        WImporte = WVector(Ciclo, 2)
        Select Case WTipopro
            Case "DY", "DS", "DQ"
                WArticulo = Left$(WProducto, 3) + Right$(WProducto, 7)
                XParam = "'" + WArticulo + "','" _
                             + WImporte + "','" _
                             + WDate + "'"
                spArticulo = "ModificaArticuloVenta " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            Case Else
                WTerminado = WProducto
                WDate = Date$
                XParam = "'" + WTerminado + "','" _
                             + WImporte + "','" _
                             + WDate + "'"
                spTerminado = "ModificaTerminadoPedido " + XParam
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        End Select
    Next Ciclo
    
    Call Cancela_click

    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    PrgProcPedPen.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub


Sub Form_Load()
End Sub

