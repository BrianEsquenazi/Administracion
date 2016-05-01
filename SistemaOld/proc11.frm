VERSION 5.00
Begin VB.Form PrgProc11 
   AutoRedraw      =   -1  'True
   Caption         =   "Generacion de Nk y RE"
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
Attribute VB_Name = "PrgProc11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WTerminado As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim XParam As String
Private Vector(10000) As String

Sub Cancelar_Click()
    PrgProc11.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Aceptar_Click()
        
    On Error GoTo WError
        
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
                
                If Left$(rstTerminado!Codigo, 2) = "PT" Then
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
    
        WTerminado = Vector(Da)
        
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WDescripcion = rstTerminado!Descripcion
            WLinea = rstTerminado!Linea
            WUnidad = rstTerminado!Unidad
            WEnvase1 = rstTerminado!Envase1
            WEnvase2 = rstTerminado!Envase2
            WEnvase3 = rstTerminado!Envase3
            WEnvase4 = rstTerminado!Envase4
            WEnvase5 = rstTerminado!Envase5
            WEnvase6 = rstTerminado!Envase6
            rstTerminado.Close
        End If
        
        WNk = "NK" + Right$(WTerminado, 10)
        
        spTerminado = "ConsultaTerminado " + "'" + WNk + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount = 0 Then
        
            rstTerminado.Close
    
            WCodigo = WNk
            WDescripcion = WDescripcion
            WLinea = Str$(WLinea)
            WUnidad = WUnidad
            WInicial = ""
            WEntradas = ""
            WSalidas = ""
            WMinimo = ""
            WDeposito = ""
            WPedido = ""
            WEnvase1 = Str$(WEnvase1)
            WEnvase2 = Str$(WEnvase2)
            WEnvase3 = Str$(WEnvase3)
            WEnvase4 = Str$(WEnvase4)
            WEnvase5 = Str$(WEnvase5)
            WEnvase6 = Str$(WEnvase6)
            WProceso = ""
            WCosto = ""
            WFactor = ""
            WDate = ""
            WImpreadi = ""
            WIntervencion = ""
            WClase = ""
            WNaciones = ""
            WEmbalaje = ""
            WVersion = ""
            WFechaVersion = "  /  /    "
            
            XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + WLinea + "','" _
                         + WUnidad + "','" _
                         + WInicial + "','" + WEntradas + "','" _
                         + WSalidas + "','" + WMinimo + "','" _
                         + WDeposito + "','" + WPedido + "','" _
                         + WEnvase1 + "','" + WEnvase2 + "','" _
                         + WEnvase3 + "','" + WEnvase4 + "','" _
                         + WEnvase5 + "','" + WEnvase6 + "','" _
                         + WProceso + "','" _
                         + WCosto + "','" _
                         + WFactor + "','" _
                         + WDate + "','" _
                         + WImpreadi + "','" _
                         + WClase + "','" _
                         + WIntervencion + "','" _
                         + WNaciones + "','" _
                         + WEmbalaje + "','" _
                         + WVersion + "','" _
                         + WFechaVersion + "'"
            
            Set rstTerminado = db.OpenRecordset("AltaTerminado " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        Rem da de alta el Re
        
        WRe = "RE" + Right$(WTerminado, 10)
        
        spTerminado = "ConsultaTerminado " + "'" + WRe + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount = 0 Then
        
            rstTerminado.Close
        
            WCodigo = WRe
            WDescripcion = WDescripcion
            WLinea = Str$(WLinea)
            WUnidad = WUnidad
            WInicial = ""
            WEntradas = ""
            WSalidas = ""
            WMinimo = ""
            WDeposito = ""
            WPedido = ""
            WEnvase1 = Str$(WEnvase1)
            WEnvase2 = Str$(WEnvase2)
            WEnvase3 = Str$(WEnvase3)
            WEnvase4 = Str$(WEnvase4)
            WEnvase5 = Str$(WEnvase5)
            WEnvase6 = Str$(WEnvase6)
            WProceso = ""
            WCosto = ""
            WFactor = ""
            WDate = ""
            WImpreadi = ""
            WIntervencion = ""
            WClase = ""
            WNaciones = ""
            WEmbalaje = ""
            WVersion = ""
            WFechaVersion = "  /  /    "
            
            XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + WLinea + "','" _
                         + WUnidad + "','" _
                         + WInicial + "','" + WEntradas + "','" _
                         + WSalidas + "','" + WMinimo + "','" _
                         + WDeposito + "','" + WPedido + "','" _
                         + WEnvase1 + "','" + WEnvase2 + "','" _
                         + WEnvase3 + "','" + WEnvase4 + "','" _
                         + WEnvase5 + "','" + WEnvase6 + "','" _
                         + WProceso + "','" _
                         + WCosto + "','" _
                         + WFactor + "','" _
                         + WDate + "','" _
                         + WImpreadi + "','" _
                         + WClase + "','" _
                         + WIntervencion + "','" _
                         + WNaciones + "','" _
                         + WEmbalaje + "','" _
                         + WVersion + "','" _
                         + WFechaVersion + "'"
                                         
            Set rstTerminado = db.OpenRecordset("AltaTerminado " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
    Next Da
        
    Call Cancelar_Click

    Exit Sub

WError:

    Resume Next

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
            PrgProc11.Caption = "Generacion de NK y RE :  " + !Nombre
        End If
    End With

End Sub
