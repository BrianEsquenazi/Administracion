VERSION 5.00
Begin VB.Form PrgProcesoFecha 
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
Attribute VB_Name = "PrgProcesoFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Trabajo(100000, 2) As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim XParam As String

Private Sub Cancelar_Click()
    PrgProcesoFecha.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()

    Erase Trabajo
    LugarTrabajo = 0
    Pasa = 0

    spLaudo = "ListaLaudoTotal "
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
    
        With rstLaudo
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte = !Laudo
                    LugarTrabajo = LugarTrabajo + 1
                    Trabajo(LugarTrabajo, 1) = !Laudo
                    Trabajo(LugarTrabajo, 2) = !Fecha
                End If
                
                If Corte <> !Laudo Then
                    Corte = !Laudo
                    LugarTrabajo = LugarTrabajo + 1
                    Trabajo(LugarTrabajo, 1) = !Laudo
                    Trabajo(LugarTrabajo, 2) = !Fecha
                End If
                
                .MoveNext
                
                If .EOF = True Then
                    Exit Do
                End If
                
            Loop
            End If
        End With
        
        rstLaudo.Close
        
    End If
    
    For Ciclo = 1 To LugarTrabajo
    
        WLaudo = Trabajo(Ciclo, 1)
        WFecha = Trabajo(Ciclo, 2)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    Erase Trabajo
    LugarTrabajo = 0
    Pasa = 0

    spHoja = "ListaHojaTotal "
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
    
            .MoveFirst
            
            If .NoMatch = False Then
            Do
            
                If .EOF = True Then
                    Exit Do
                End If
                
                If Pasa = 0 Then
                    Pasa = 1
                    Corte = !Hoja
                    LugarTrabajo = LugarTrabajo + 1
                    Trabajo(LugarTrabajo, 1) = !Hoja
                    Trabajo(LugarTrabajo, 2) = !Fecha
                End If
                
                If Corte <> !Hoja Then
                    Corte = !Hoja
                    LugarTrabajo = LugarTrabajo + 1
                    Trabajo(LugarTrabajo, 1) = !Hoja
                    Trabajo(LugarTrabajo, 2) = !Fecha
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
    
    For Ciclo = 1 To LugarTrabajo
    
        WHoja = Trabajo(Ciclo, 1)
        WFecha = Trabajo(Ciclo, 2)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        
        XParam = "'" + WHoja + "','" _
                     + WFechaord + "'"
        Set rstHoja = db.OpenRecordset("ModificaHojaFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    
    
    
    Call Cancelar_Click
    
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

