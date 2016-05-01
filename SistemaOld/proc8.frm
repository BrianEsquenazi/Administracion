VERSION 5.00
Begin VB.Form PrgProc8 
   AutoRedraw      =   -1  'True
   Caption         =   "Invantario PT"
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
Attribute VB_Name = "PrgProc8"
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
Private Wcampo1 As String
Private Wcampo2 As String
Private Wcampo3 As String
Private Wcampo4 As Double

Private Sub Cancelar_Click()

    With rstTerminado
        .Close
    End With
    With rstPt
        .Close
    End With
    
    DbsVentas.Close
    
    PrgProc8.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Open "lpt1" For Output As #1

    With rstPt
    
            .Index = "clave"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                Wcampo1 = !Campo1
                Wcampo2 = Str$(!Campo2)
                If !Campo3 <> Null Then
                    If Val(!Campo3) = 0 Then
                        Wcampo3 = "100"
                            Else
                        Wcampo3 = Str$(Val(!Campo3))
                    End If
                        Else
                    Wcampo3 = "100"
                End If
                Wcampo4 = !Campo4
                
                Call Ceros(Wcampo2, 5)
                Call Ceros(Wcampo3, 3)
                
                WClave = Wcampo1 + "-" + Wcampo2 + "-" + Wcampo3
                
                With rstTerminado
    
                    .Index = "Codigo"
                    .Seek "=", WClave
            
                    If .NoMatch = False Then
                        .Edit
                        !Inicial = !Inicial + Wcampo4
                        .Update
                            Else
                         Print #1, WClave, Wcampo4
                    End If
                    
                End With
                                        
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    Close #1
    
    Call Cancelar_Click

End Sub

Private Sub Form_Activate()
    OPEN_FILE_TERMINADO
    OPEN_FILE_Pt
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgProc8.Caption = "Inventario de P.T. :  " + !Nombre
        End If
    End With

End Sub
