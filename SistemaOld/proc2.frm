VERSION 5.00
Begin VB.Form PrgProc2 
   AutoRedraw      =   -1  'True
   Caption         =   "Reprocesos de Productos Terminados"
   ClientHeight    =   7170
   ClientLeft      =   225
   ClientTop       =   975
   ClientWidth     =   11655
   LinkTopic       =   "Form2"
   ScaleHeight     =   7170
   ScaleWidth      =   11655
   Begin VB.ComboBox Tipo 
      Height          =   315
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
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
Attribute VB_Name = "PrgProc2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Private WEntradas As Double
Private WSalidas As Double
Private Vector(10000) As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
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
Dim rstConsig As Recordset
Dim spConsig As String
Dim XParam As String

Private Sub Cancelar_Click()
    PrgProc2.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Erase Vector
    Renglon = 0
        
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
    
    With rstTerminado
        .MoveFirst
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Left$(rstTerminado!Codigo, 2) = "PT" Or Tipo.ListIndex = 1 Then
                    Renglon = Renglon + 1
                    Vector(Renglon) = rstTerminado!Codigo
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
    End With
    rstTerminado.Close
    
    End If
    
    For Da = 1 To Renglon
    
        WEntradas = 0
        WSalidas = 0
        WTerminado = Vector(Da)
        XCodigo = Vector(Da)
        XDate = Date$
        
        Call calcula_datos
        
        XEntradas = Str$(WEntradas)
        XSalidas = Str$(WSalidas)
        
        XParam = "'" + XCodigo + "','" _
                + XEntradas + "','" _
                + XSalidas + "','" _
                + XDate + "'"
                                           
        spTerminado = "ModificaTerminadoMovimientos " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        Rem End If
        
    Next Da
    
    Call Cancelar_Click

End Sub

Private Sub calcula_datos()

    spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
        WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
        rstTerminado.Close
    End If

    Rem PROCESA LAS ESTADISTICAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spEstadistica = "ListaEstadisticaRepro" + XParam
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
    
        With rstEstadistica
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstEstadistica!Marca = "X" Then
                
                    Else
                
                If Val(rstEstadistica!Tipo) = 1 Then
                    WSalidas = WSalidas + rstEstadistica!Cantidad
                        Else
                    WEntradas = WEntradas + Abs(rstEstadistica!Cantidad)
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
            
        End With
        
        rstEstadistica.Close
        
    End If
    
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro1" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                XFec = Right$(rstHoja!Fecha, 4) + Mid$(rstHoja!Fecha, 4, 2) + Left$(rstHoja!Fecha, 2)
                If rstHoja!Marca = "X" Or XFec < WOrdFechaCierre Then
                
                    Else
                            
                If rstHoja!Tipo = "T" Then
                
                    WSalidas = WSalidas + rstHoja!Cantidad
                
                End If
                
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            
            End If
            
        End With
    End If
    
    Rem PROCESA LAS HOJAS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spHoja = "ListaHojaRepro2" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
            
                If rstHoja!Marca = "X" And rstHoja!Saldo = 0 Then
                
                    Else
                
                If Val(rstHoja!Renglon) = 1 And rstHoja!Real <> 0 Then
                
                    WEntradas = WEntradas + rstHoja!Real
                    
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
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovvar = "ListaMovvarRepro" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovvar.RecordCount > 0 Then
    
        With rstMovvar
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovvar!Marca = "X" Then
                
                        Else
                
                If rstMovvar!Tipo = "T" Then
                
                    If rstMovvar!Movi = "E" Then
                        WEntradas = WEntradas + rstMovvar!Cantidad
                            Else
                        WSalidas = WSalidas + rstMovvar!Cantidad
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
        
        rstMovvar.Close
    End If
    
    
    
    Rem PROCESA LOS MOVIMIENTOS VARIOS
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovguia = "ListaMovguiaRepro" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovguia!Marca = "X" And rstMovguia!Saldo = 0 Then
                
                        Else
                
                If rstMovguia!Tipo = "T" Then
                
                    If rstMovguia!Movi = "E" Then
                        WEntradas = WEntradas + rstMovguia!Cantidad
                            Else
                        WSalidas = WSalidas + rstMovguia!Cantidad
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
        
        rstMovguia.Close
    End If
    
    
    Rem PROCESA LOS MOVIMIENTOS DE LABORATORIO
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spMovlab = "ListaMovlabRepro" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovlab.RecordCount > 0 Then
    
        With rstMovlab
    
            .MoveFirst
            
            If .NoMatch = False Then
            
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstMovlab!Marca = "X" Then
                
                    Else
                
                If rstMovlab!Tipo = "T" Then
                
                    If rstMovlab!Movi = "E" Then
                        WEntradas = WEntradas + rstMovlab!Cantidad
                                Else
                        WSalidas = WSalidas + rstMovlab!Cantidad
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
        
        rstMovlab.Close
    End If
    
    XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
    spConsig = "ListaConsigRepro" + XParam
    Set rstConsig = db.OpenRecordset(spConsig, dbOpenSnapshot, dbSQLPassThrough)
    If rstConsig.RecordCount > 0 Then
    
        With rstConsig
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If rstConsig!Marca <> "X" Then
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WSalidas = WSalidas + WCantidad
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        rstConsig.Close
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgProc2.Caption = "Reproceso de Productos Terminados :  " + !Nombre
        End If
    End With
    
    Tipo.Clear
    
    Tipo.AddItem "PT"
    Tipo.AddItem "Total"
    
    Tipo.ListIndex = 0

End Sub
