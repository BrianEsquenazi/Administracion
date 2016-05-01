VERSION 5.00
Begin VB.Form PrgProc6 
   AutoRedraw      =   -1  'True
   Caption         =   "Cierre de Stock de materias Primas y Producto Terminado"
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
Attribute VB_Name = "PrgProc6"
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

Private Sub Cancelar_Click()
    With rstHoja
        .Close
    End With
    With rstMovvar
        .Close
    End With
    With rstLaudo
        .Close
    End With
    With rstMovlab
        .Close
    End With
    With rstEstadistica
        .Close
    End With
    With rstEmpresa
        .Close
    End With
    With rstAuxiliar
        .Close
    End With
    DbsVentas.Close
    DbsCotiza.Close
    PrgProc3.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    With rstLaudo
    
            .Index = "clave"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                .Edit
                !Marca = "X"
                .Update
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
    End With
    
    With rstHoja
    
            .Index = "Articulo"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                .Edit
                !Marca = "X"
                .Update
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
        
    End With
    
    With rstMovvar
    
            .Index = "Articulo"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                .Edit
                !Marca = "X"
                .Update
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With
    
    With rstMovlab
    
            .Index = "Articulo"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                .Edit
                !Marca = "X"
                .Update
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With
    
    With rstEstadistica
    
            .Index = "Articulo"
            .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                .Edit
                !Marca = "X"
                .Update
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With

End Sub
