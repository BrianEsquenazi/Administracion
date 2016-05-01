VERSION 5.00
Begin VB.Form PrgProchoja 
   AutoRedraw      =   -1  'True
   Caption         =   "Verificacion de hojas de produccion"
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
Attribute VB_Name = "PrgProchoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String

Private Sub Cancelar_Click()
    PrgProchoja.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Open "lpt1" For Output As #1
    Rem Open "PROCHOJA.TXT" For Output As #1
    
    spHoja = "ListaHojaTotal"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!fechaIng, 4) + Mid$(rstHoja!fechaIng, 4, 2) + Left$(rstHoja!fechaIng, 2)
                If rstHoja!Marca = "X" Then
                
                        Else
                    
                    If rstHoja!Articulo <> "AA-000-100" And rstHoja!Articulo <> "ZC-001-100" Then
                    
                        Arti1 = rstHoja!Terminado
                        Arti2 = rstHoja!Articulo
                        Hoja = rstHoja!Hoja
                        Canti1 = rstHoja!Cantidad
                        Canti2 = rstHoja!Canti1 + rstHoja!Canti2 + rstHoja!Canti3
                
                        If Canti1 <> Canti2 Then
                            Print #1, Arti1, Arti2, Hoja, Canti1, Canti2
                        End If
                        
                    End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
        
        rstHoja.Close
        
    End If
    
    Close #1
    
    Call Cancelar_Click
    
End Sub


