VERSION 5.00
Begin VB.Form PrgVeriForm 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de Formulas"
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
Attribute VB_Name = "PrgVeriForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(10000, 3) As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstProveedor As Recordset
Dim spProveedor As String
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
Dim XParam As String

Private Sub Cancelar_Click()

    PrgVeriForm.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    Pasa = 0
    Open "Compo.txt" For Output As #1

    spComposicion = "ListaComposicion"
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
        
    With rstComposicion

            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Pasa = 0 Then
                    Pasa = 1
                    corte = rstComposicion!Terminado
                    Renglon = 0
                    Suma = 0
                End If
                
                If corte <> rstComposicion!Terminado Then
                    If Renglon <> Mayor Then
                    Rem If Suma < 1 Or Renglon <> Mayor Then
                        If Left$(corte, 2) = "PT" Then
                            Print #1, corte, Suma, Renglon, Mayor
                        End If
                    End If
                    corte = rstComposicion!Terminado
                    Renglon = 0
                    Suma = 0
                End If

                Renglon = Renglon + 1
                Mayor = Val(rstComposicion!Renglon)
                Suma = Suma + rstComposicion!Cantidad
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


