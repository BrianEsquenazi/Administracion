VERSION 5.00
Begin VB.Form PrgProchojaEspecif 
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
Attribute VB_Name = "PrgProchojaEspecif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim Vector(10000, 3) As String

Private Sub Cancelar_Click()
    PrgProchojaEspecif.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Erase Vector
    Lugar = 0
    
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
                
                
                aa = rstHoja!Clave
                
                If rstHoja!Fechaord >= "20051021" Then
                
                    Entra = "S"
                    For Ciclo = 1 To Lugar
                        If Val(Vector(Ciclo, 1)) = rstHoja!Hoja Then
                            Entra = "N"
                            Exit For
                        End If
                    Next Ciclo
                    
                    If Entra = "S" Then
                        Lugar = Lugar + 1
                        Vector(Lugar, 1) = Str$(rstHoja!Hoja)
                        Vector(Lugar, 2) = IIf(IsNull(rstHoja!VersionIII), "", rstHoja!VersionIII)
                        Vector(Lugar, 3) = rstHoja!Fecha
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
    
    Stop
    
    aa = Lugar
    
    For Ciclo = 1 To Lugar
    
        ZHoja = Vector(Lugar, 1)
        ZVersion = Vector(Lugar, 2)
        ZFecha = Vector(Lugar, 3)
        
    
    
    
End Sub


