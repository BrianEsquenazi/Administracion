VERSION 5.00
Begin VB.Form Prglee1 
   Caption         =   "Generacion de traspaso de datos"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.Frame Frame2 
      Caption         =   "Control de Grabacion"
      Height          =   1815
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cerrar"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Prglee1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLinea As String
Private WEmpresa As String

Private Sub Acepta_Click()
 
    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    Prglee1.Hide
    Unload Me
    Menu.SetFocus
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    'ensayos
        
    coderr = 0
    
    DADA = "c:\prueba\labora\" + WEmpresa + "ensa.txt"
    Open "c:\prueba\labora\" + WEmpresa + "ensa.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
            
                WEnsayo = Mid$(WLinea, 5, 4)
                WDescripcion = Mid$(WLinea, 9, 40)
                    
                With rstEnsayos
                        .Index = "Codigo"
                        .Seek "=", Val(WEnsayo)
                        If .NoMatch Then
                            .AddNew
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = WEnsayo
                            !Descripcion = WDescripcion
                            !Wdate = Date$
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
    
    'PRUEBAS DE MATERIAS PRIMAS
        
    coderr = 0
    
    Open "c:\prueba\labora\" + WEmpresa + "inp.txt" For Input As #1
    
    Do
                Line Input #1, WLinea

                WPrueba = Mid$(WLinea, 5, 1) + Mid$(WLinea, 7, 5)
                WProducto = Mid$(WLinea, 12, 2) + "-" + Mid$(WLinea, 14, 3) + "-" + Mid$(WLinea, 17, 3)
                WFecha = Mid$(WLinea, 23, 2) + "/" + Mid$(WLinea, 25, 2) + "/19" + Mid$(WLinea, 27, 2)
                WOrden = Mid$(WLinea, 29, 6)
                WValor1 = Mid$(WLinea, 35, 15)
                WValor2 = Mid$(WLinea, 50, 15)
                WValor3 = Mid$(WLinea, 65, 15)
                WValor4 = Mid$(WLinea, 80, 15)
                WValor5 = Mid$(WLinea, 95, 15)
                WValor6 = Mid$(WLinea, 110, 15)
                WValor7 = Mid$(WLinea, 125, 15)
                WValor8 = Mid$(WLinea, 140, 15)
                WValor9 = Mid$(WLinea, 155, 15)
                WValor10 = Mid$(WLinea, 170, 15)
                WEnsayo = Mid$(WLinea, 185, 40)
                WAspecto = Mid$(WLinea, 225, 40)
                WObservaciones = Left$(Mid$(WLinea, 265, 120), 50)
                WConfecciono = Mid$(WLinea, 385, 40)
                WLiberada = 0
                WDevuelta = 0
                WLote = 0
                WRechazo = 0
                WNueva = "N"
                
                With rstPrueba
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Orden = WOrden
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !Nueva = WNueva
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Orden = WOrden
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Liberada = WLiberada
                            !Devuelta = WDevuelta
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !Nueva = WNueva
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
    
    
        'PRUEBAS DE productos terminados
        
    coderr = 0
    
    Open "c:\prueba\labora\" + WEmpresa + "ing.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                
                WPrueba = Mid$(WLinea, 5, 1) + Mid$(WLinea, 7, 5)
                WProducto = Mid$(WLinea, 12, 2) + "-" + Mid$(WLinea, 14, 5) + "-" + Mid$(WLinea, 19, 3)
                WFecha = Mid$(WLinea, 22, 2) + "/" + Mid$(WLinea, 24, 2) + "/19" + Mid$(WLinea, 26, 2)
                WValor1 = Mid$(WLinea, 34, 15)
                WValor2 = Mid$(WLinea, 49, 15)
                WValor3 = Mid$(WLinea, 64, 15)
                WValor4 = Mid$(WLinea, 79, 15)
                WValor5 = Mid$(WLinea, 94, 15)
                WValor6 = Mid$(WLinea, 109, 15)
                WValor7 = Mid$(WLinea, 124, 15)
                WValor8 = Mid$(WLinea, 139, 15)
                WValor9 = Mid$(WLinea, 154, 15)
                WValor10 = Mid$(WLinea, 169, 15)
                WEnsayo = Mid$(WLinea, 184, 40)
                WAspecto = Mid$(WLinea, 224, 40)
                WObservaciones = Left$(Mid$(WLinea, 264, 120), 50)
                WConfecciono = Mid$(WLinea, 384, 40)
                WLote = 0
                WRechazo = 0
                
                With rstPrueter
                        .Index = "Prueba"
                        .Seek "=", WPrueba
                        If .NoMatch Then
                            .AddNew
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Prueba = WPrueba
                            !Producto = WProducto
                            !Fecha = WFecha
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            !Ensayo = WEnsayo
                            !Aspecto = WAspecto
                            !Observaciones = WObservaciones
                            !Confecciono = WConfecciono
                            !Lote = WLote
                            !Rechazo = WRechazo
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
    

    
    'especificaciones de p.t
    
    coderr = 0
    
    Open "c:\prueba\labora\" + WEmpresa + "pru.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                If Mid$(WLinea, 5, 2) = "PT" Then
        
                WProducto = Mid$(WLinea, 5, 2) + "-" + Mid$(WLinea, 7, 5) + "-" + Mid$(WLinea, 13, 3)
                WEnsayo1 = Val(Mid$(WLinea, 16, 4))
                WEnsayo2 = Val(Mid$(WLinea, 35, 4))
                WEnsayo3 = Val(Mid$(WLinea, 54, 4))
                WEnsayo4 = Val(Mid$(WLinea, 73, 4))
                WEnsayo5 = Val(Mid$(WLinea, 92, 4))
                WEnsayo6 = Val(Mid$(WLinea, 111, 4))
                WEnsayo7 = Val(Mid$(WLinea, 130, 4))
                WEnsayo8 = 0
                Rem Mid$(WLinea, 149, 4)
                WEnsayo9 = 0
                Rem Mid$(WLinea, 168, 4)
                WEnsayo10 = 0
                Rem Mid$(WLinea, 187, 4)
                WValor1 = Mid$(WLinea, 20, 15)
                WValor2 = Mid$(WLinea, 39, 15)
                WValor3 = Mid$(WLinea, 58, 15)
                WValor4 = Mid$(WLinea, 77, 15)
                WValor5 = Mid$(WLinea, 96, 15)
                WValor6 = Mid$(WLinea, 115, 15)
                WValor7 = Mid$(WLinea, 134, 15)
                WValor8 = ""
                Rem Mid$(WLinea, 153, 15)
                WValor9 = ""
                Rem Mid$(WLinea, 172, 15)
                WValor10 = ""
                Rem Mid$(WLinea, 191, 15)
                
                With rstEspecif
                        .Index = "Producto"
                        .Seek "=", WProducto
                        If .NoMatch Then
                            .AddNew
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                End If
                If EOF(1) Then Exit Do
    Loop
    
    Close #1
    


    'especificaciones de m.p.
        
    coderr = 0
    
    Open "c:\prueba\labora\" + WEmpresa + "pru.txt" For Input As #1
    
    Do
                Line Input #1, WLinea
                
                If Mid$(WLinea, 5, 2) <> "PT" Then
        
                WProducto = Mid$(WLinea, 5, 2) + "-" + Mid$(WLinea, 7, 3) + "-" + Mid$(WLinea, 10, 3)
                WEnsayo1 = Val(Mid$(WLinea, 16, 4))
                WEnsayo2 = Val(Mid$(WLinea, 35, 4))
                WEnsayo3 = Val(Mid$(WLinea, 54, 4))
                WEnsayo4 = Val(Mid$(WLinea, 73, 4))
                WEnsayo5 = Val(Mid$(WLinea, 92, 4))
                WEnsayo6 = Val(Mid$(WLinea, 111, 4))
                WEnsayo7 = Val(Mid$(WLinea, 130, 4))
                WEnsayo8 = 0
                Rem Mid$(WLinea, 149, 4)
                WEnsayo9 = 0
                Rem Mid$(WLinea, 168, 4)
                WEnsayo10 = 0
                Rem Mid$(WLinea, 187, 4)
                WValor1 = Mid$(WLinea, 20, 15)
                WValor2 = Mid$(WLinea, 39, 15)
                WValor3 = Mid$(WLinea, 58, 15)
                WValor4 = Mid$(WLinea, 77, 15)
                WValor5 = Mid$(WLinea, 96, 15)
                WValor6 = Mid$(WLinea, 115, 15)
                WValor7 = Mid$(WLinea, 134, 15)
                WValor8 = ""
                Rem Mid$(WLinea, 153, 15)
                WValor9 = ""
                Rem Mid$(WLinea, 172, 15)
                WValor10 = ""
                Rem Mid$(WLinea, 191, 15)
        
                With rstEspecificaciones
                        .Index = "Producto"
                        .Seek "=", WProducto
                        If .NoMatch Then
                            .AddNew
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Producto = WProducto
                            !Ensayo1 = WEnsayo1
                            !Ensayo2 = WEnsayo2
                            !Ensayo3 = WEnsayo3
                            !Ensayo4 = WEnsayo4
                            !Ensayo5 = WEnsayo5
                            !Ensayo6 = WEnsayo6
                            !Ensayo7 = WEnsayo7
                            !Ensayo8 = WEnsayo8
                            !Ensayo9 = WEnsayo9
                            !Ensayo10 = WEnsayo10
                            !Valor1 = WValor1
                            !valor2 = WValor2
                            !Valor3 = WValor3
                            !valor4 = WValor4
                            !valor5 = WValor5
                            !valor6 = WValor6
                            !valor7 = WValor7
                            !valor8 = WValor8
                            !valor9 = WValor9
                            !valor10 = WValor10
                            .Update
                            .Bookmark = .LastModified
                        End If
                End With
                End If
                If EOF(1) Then Exit Do
    Loop
    
    Close #1




    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Exit Sub
    
    
Error:
Stop
     coderr = Err
     Resume Next
     
End Sub




