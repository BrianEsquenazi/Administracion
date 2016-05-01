VERSION 5.00
Begin VB.Form PrgRecpel2 
   Caption         =   "Recepcion de Informacion"
   ClientHeight    =   4620
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   6390
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   4620
   ScaleWidth      =   6390
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SISTEMA EN PROCESO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   4335
   End
End
Attribute VB_Name = "PrgRecpel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstPrueart As Recordset
Dim spPrueart As String
Dim rstPrueter As Recordset
Dim spPrueter As String
Dim rstEspecif As Recordset
Dim spEspecif As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim rstMovlab As Recordset
Dim spmovlab As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCtacte As Recordset
Dim spCtacte As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstDesccomp As Recordset
Dim spDesccomp As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim XParam As String
Dim A1 As String
Dim A2 As String
Dim Auxi As String
Dim Auxi1 As String
Dim WTipomov As String

Private Sub Acepta_Click()

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    WEmpresa = "0002"
    txtOdbc = "Empresa02"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso
    
    WEmpresa = "0003"
    txtOdbc = "Empresa03"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso1
    
    WEmpresa = "0004"
    txtOdbc = "Empresa04"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso1
    
    WEmpresa = "0005"
    txtOdbc = "Empresa05"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso1
    
    WEmpresa = "0006"
    txtOdbc = "Empresa06"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    Call Proceso1

    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgRecpel2.Hide
    Unload Me
    End
End Sub


Private Sub Proceso()

    On Error GoTo Error
    
    Rem OPEN_FILE_WENSAYOS
    Rem OPEN_FILE_WESPECIFICACIONES
    Rem OPEN_FILE_WESPECIF
    OPEN_FILE_WPRUEBA
    OPEN_FILE_WPrueTer
    OPEN_FILE_WMovlab
    OPEN_FILE_WHoja
    OPEN_FILE_WInforme
    OPEN_FILE_WLAUDO
    OPEN_FILE_WMovvar
    OPEN_FILE_WMovguia
    OPEN_FILE_WComposicion
    OPEN_FILE_WTERMINADO
    OPEN_FILE_WEstadistica
    OPEN_FILE_WArticulo
    OPEN_FILE_WCtacte
    OPEN_FILE_WProveedor
    
    'ensayos

    Rem coderr = 0
    Rem With rstWEnsayos
    Rem         .Index = "Codigo"
    Rem         .MoveFirst
    Rem         Do
    Rem
    Rem             WEnsayo = Str$(!Codigo)
    Rem             WDescripcion = !Descripcion
    Rem             WDate = ""
    Rem
    Rem             spEnsayo = "ConsultaEnsayos " + "'" + WEnsayo + "'"
    Rem             Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstEnsayo.RecordCount > 0 Then
    Rem
    Rem                 rstEnsayo.Close
    Rem                 XParam = "'" + WEnsayo + "','" _
    rem                      + WDescripcion + "','" _
    rem                      + WDate + "'"
    Rem
    Rem                 Set rstEnsayo = db.OpenRecordset("ModificaEnsayos " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem                     Else
    Rem
    Rem                 XParam = "'" + WEnsayo + "','" _
    rem                      + WDescripcion + "','" _
    rem                      + WDate + "'"
    Rem
    Rem                 Set rstEnsayo = db.OpenRecordset("AltaEnsayos " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With
    
    'PRUEBAS DE MATERIAS PRIMAS
    
    coderr = 0
    With rstWPrueba
            .Index = "Prueba"
            .MoveFirst
            Do
            
                WPrueba = !Prueba
                WProducto = !Producto
                WFecha = !Fecha
                WOrden = !Orden
                WValor1 = !Valor1
                Wvalor2 = !valor2
                WValor3 = !Valor3
                Wvalor4 = !valor4
                Wvalor5 = !valor5
                Wvalor6 = !valor6
                Wvalor7 = !valor7
                Wvalor8 = !valor8
                Wvalor9 = !valor9
                Wvalor10 = !valor10
                WEnsayo = !Ensayo
                WAspecto = !Aspecto
                WObservaciones = !Observaciones
                WObserva2 = ""
                WConfecciono = !Confecciono
                WLiberada = Str$(!Liberada)
                WDevuelta = Str$(!Devuelta)
                WLote = Str$(!Lote)
                WRechazo = Str$(!Rechazo)
                WNueva = !Nueva
                WFechaord = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                WDate = ""
                
                spPrueart = "ConsultaPrueart " + "'" + WPrueba + "'"
                Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrueart.RecordCount > 0 Then
                
                    rstPrueart.Close
                    XParam = "'" + WPrueba + "','" _
                                + WProducto + "','" _
                                + WFecha + "','" _
                                + WOrden + "','" _
                                + WValor1 + "','" _
                                + Wvalor2 + "','" _
                                + WValor3 + "','" _
                                + Wvalor4 + "','" _
                                + Wvalor5 + "','" _
                                + Wvalor6 + "','" _
                                + Wvalor7 + "','" _
                                + Wvalor8 + "','" _
                                + Wvalor9 + "','" _
                                + Wvalor10 + "','" _
                                + WEnsayo + "','" _
                                + WAspecto + "','" _
                                + WObservaciones + "','" _
                                + WObserva2 + "','" _
                                + WConfecciono + "','" _
                                + WLiberada + "','" _
                                + WDevuelta + "','" _
                                + WLote + "','" _
                                + WRechazo + "','" _
                                + WNueva + "','" + WFechaord + "','" _
                                + WDate + "'"
                                
                    Set rstPrueart = db.OpenRecordset("ModificaPrueart " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XParam = "'" + WPrueba + "','" _
                                + WProducto + "','" _
                                + WFecha + "','" _
                                + WOrden + "','" _
                                + WValor1 + "','" _
                                + Wvalor2 + "','" _
                                + WValor3 + "','" _
                                + Wvalor4 + "','" _
                                + Wvalor5 + "','" _
                                + Wvalor6 + "','" _
                                + Wvalor7 + "','" _
                                + Wvalor8 + "','" _
                                + Wvalor9 + "','" _
                                + Wvalor10 + "','" _
                                + WEnsayo + "','" _
                                + WAspecto + "','" _
                                + WObservaciones + "','" _
                                + WObserva2 + "','" _
                                + WConfecciono + "','" _
                                + WLiberada + "','" _
                                + WDevuelta + "','" _
                                + WLote + "','" _
                                + WRechazo + "','" _
                                + WNueva + "','" + WFechaord + "','" _
                                + WDate + "'"
        
                    Set rstPrueart = db.OpenRecordset("AltaPrueart " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'PRUEBAS DE productos terminados
        
    coderr = 0
    With rstWPrueter
            .Index = "Prueba"
            .MoveFirst
            Do
            
                WPrueba = !Prueba
                WProducto = !Producto
                WFecha = !Fecha
                WValor1 = !Valor1
                Wvalor2 = !valor2
                WValor3 = !Valor3
                Wvalor4 = !valor4
                Wvalor5 = !valor5
                Wvalor6 = !valor6
                Wvalor7 = !valor7
                Wvalor8 = !valor8
                Wvalor9 = !valor9
                Wvalor10 = !valor10
                WEnsayo = !Ensayo
                WAspecto = !Aspecto
                WObservaciones = !Observaciones
                WConfecciono = !Confecciono
                WLiberada = Str$(!Liberada)
                WLote = Str$(!Lote)
                WRechazo = Str$(!Rechazo)
                WFechaord = Right$(!Fecha, 4) + Mid$(!Fecha, 4, 2) + Left$(!Fecha, 2)
                WDate = ""
                
                spPrueter = "ConsultaPrueter " + "'" + WPrueba + "'"
                Set rstPrueter = db.OpenRecordset(spPrueter, dbOpenSnapshot, dbSQLPassThrough)
                If rstPrueter.RecordCount > 0 Then
                    rstPrueter.Close
                
                    XParam = "'" + WPrueba + "','" _
                                + WProducto + "','" _
                                + WFecha + "','" _
                                + WValor1 + "','" _
                                + Wvalor2 + "','" _
                                + WValor3 + "','" _
                                + Wvalor4 + "','" _
                                + Wvalor5 + "','" _
                                + Wvalor6 + "','" _
                                + Wvalor7 + "','" _
                                + Wvalor8 + "','" _
                                + Wvalor9 + "','" _
                                + Wvalor10 + "','" _
                                + WEnsayo + "','" _
                                + WAspecto + "','" _
                                + WObservaciones + "','" _
                                + WConfecciono + "','" _
                                + WLiberada + "','" _
                                + WLote + "','" _
                                + WRechazo + "','" _
                                + WFechaord + "','" _
                                + WDate + "'"
                
                                
                    Set rstPrueter = db.OpenRecordset("ModificaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XParam = "'" + WPrueba + "','" _
                                + WProducto + "','" _
                                + WFecha + "','" _
                                + WValor1 + "','" _
                                + Wvalor2 + "','" _
                                + WValor3 + "','" _
                                + Wvalor4 + "','" _
                                + Wvalor5 + "','" _
                                + Wvalor6 + "','" _
                                + Wvalor7 + "','" _
                                + Wvalor8 + "','" _
                                + Wvalor9 + "','" _
                                + Wvalor10 + "','" _
                                + WEnsayo + "','" _
                                + WAspecto + "','" _
                                + WObservaciones + "','" _
                                + WConfecciono + "','" _
                                + WLiberada + "','" _
                                + WLote + "','" _
                                + WRechazo + "','" _
                                + WFechaord + "','" _
                                + WDate + "'"
        
                    Set rstPrueter = db.OpenRecordset("AltaPrueter " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    'especificaciones de p.t.
        
    Rem coderr = 0
    Rem With rstWEspecif
    Rem         .Index = "Producto"
    Rem         .MoveFirst
    Rem
    Rem         Do
    Rem
    Rem             WProducto = !Producto
    Rem             WEnsayo1 = Str$(!Ensayo1)
    Rem             WEnsayo2 = Str$(!Ensayo2)
    Rem             WEnsayo3 = Str$(!Ensayo3)
    Rem             WEnsayo4 = Str$(!Ensayo4)
    Rem             WEnsayo5 = Str$(!Ensayo5)
    Rem             WEnsayo6 = Str$(!Ensayo6)
    Rem             WEnsayo7 = Str$(!Ensayo7)
    Rem             WEnsayo8 = Str$(!Ensayo8)
    Rem             WEnsayo9 = Str$(!Ensayo9)
    Rem             WEnsayo10 = Str$(!Ensayo10)
    Rem             WValor1 = !Valor1
    Rem             Wvalor2 = !valor2
    Rem             WValor3 = !Valor3
    Rem             Wvalor4 = !valor4
    Rem             Wvalor5 = !valor5
    Rem             Wvalor6 = !valor6
    Rem             Wvalor7 = !valor7
    Rem             Wvalor8 = !valor8
    Rem             Wvalor9 = !valor9
    Rem             Wvalor10 = !valor10
    Rem             WDate = ""
    Rem
    Rem             spEspecif = "ConsultaEspecif " + "'" + WProducto + "'"
    Rem             Set rstEspecif = db.OpenRecordset(spEspecif, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstEspecif.RecordCount > 0 Then
    Rem                 rstEspecif.Close
    Rem                 XParam = "'" + WProducto + "','" _
    rem                     + WEnsayo1 + "','" _
    rem                     + WValor1 + "','" _
    rem                     + WEnsayo2 + "','" _
    rem                     + Wvalor2 + "','" _
    rem                     + WEnsayo3 + "','" _
    rem                     + WValor3 + "','" _
    rem                     + WEnsayo4 + "','" _
    rem                     + Wvalor4 + "','" _
    rem                     + WEnsayo5 + "','" _
    rem                     + Wvalor5 + "','" _
    rem                     + WEnsayo6 + "','" _
    rem                     + Wvalor6 + "','" _
    rem                     + WEnsayo7 + "','" _
    rem                     + Wvalor7 + "','" _
    rem                     + WEnsayo8 + "','" _
    rem                     + Wvalor8 + "','" _
    rem                     + WEnsayo9 + "','" _
    rem                     + Wvalor9 + "','" _
    rem                     + WEnsayo10 + "','" _
    rem                     + Wvalor10 + "','" _
    rem                     + WDate + "'"
    Rem                 Set rstEspecif = db.OpenRecordset("ModificaEspecif" + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem                         Else
    Rem                 XParam = "'" + WProducto + "','" _
    rem                     + WEnsayo1 + "','" _
    rem                     + WValor1 + "','" _
    rem                     + WEnsayo2 + "','" _
    rem                     + Wvalor2 + "','" _
    rem                     + WEnsayo3 + "','" _
    rem                     + WValor3 + "','" _
    rem                     + WEnsayo4 + "','" _
    rem                     + Wvalor4 + "','" _
    rem                     + WEnsayo5 + "','" _
    rem                     + Wvalor5 + "','" _
    rem                     + WEnsayo6 + "','" _
    rem                     + Wvalor6 + "','" _
    rem                     + WEnsayo7 + "','" _
    rem                     + Wvalor7 + "','" _
    rem                     + WEnsayo8 + "','" _
    rem                     + Wvalor8 + "','" _
    rem                     + WEnsayo9 + "','" _
    rem                     + Wvalor9 + "','" _
    rem                     + WEnsayo10 + "','" _
    rem                     + Wvalor10 + "','" _
    rem                     + WDate + "'"
    Rem                 Set rstEspecif = db.OpenRecordset("AltaEspecif " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With


    'especificaciones de m.p.
        
    Rem coderr = 0
    Rem With rstWEspecificaciones
    Rem         .Index = "Producto"
    Rem         .MoveFirst
    Rem         Do
    Rem
    Rem             WProducto = !Producto
    Rem             WEnsayo1 = Str$(!Ensayo1)
    Rem             WEnsayo2 = Str$(!Ensayo2)
    Rem             WEnsayo3 = Str$(!Ensayo3)
    Rem             WEnsayo4 = Str$(!Ensayo4)
    Rem             WEnsayo5 = Str$(!Ensayo5)
    Rem             WEnsayo6 = Str$(!Ensayo6)
    Rem             WEnsayo7 = Str$(!Ensayo7)
    Rem             WEnsayo8 = Str$(!Ensayo8)
    Rem             WEnsayo9 = Str$(!Ensayo9)
    Rem             WEnsayo10 = Str$(!Ensayo10)
    Rem             WValor1 = !Valor1
    Rem             Wvalor2 = !valor2
    Rem             WValor3 = !Valor3
    Rem             Wvalor4 = !valor4
    Rem             Wvalor5 = !valor5
    Rem             Wvalor6 = !valor6
    Rem             Wvalor7 = !valor7
    Rem             Wvalor8 = !valor8
    Rem             Wvalor9 = !valor9
    Rem             Wvalor10 = !valor10
    Rem             WDate = ""
    Rem
    Rem             spEspecificaciones = "ConsultaEspecificaciones " + "'" + WProducto + "'"
    Rem             Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    Rem             If rstEspecificaciones.RecordCount > 0 Then
    Rem                 rstEspecificaciones.Close
    Rem                 XParam = "'" + WProducto + "','" _
    rem                     + WEnsayo1 + "','" _
    rem                     + WValor1 + "','" _
    rem                     + WEnsayo2 + "','" _
    rem                     + Wvalor2 + "','" _
    rem                     + WEnsayo3 + "','" _
    rem                     + WValor3 + "','" _
    rem                     + WEnsayo4 + "','" _
    rem                     + Wvalor4 + "','" _
    rem                     + WEnsayo5 + "','" _
    rem                     + Wvalor5 + "','" _
    rem                     + WEnsayo6 + "','" _
    rem                     + Wvalor6 + "','" _
    rem                     + WEnsayo7 + "','" _
    rem                     + Wvalor7 + "','" _
    rem                     + WEnsayo8 + "','" _
    rem                     + Wvalor8 + "','" _
    rem                     + WEnsayo9 + "','" _
    rem                     + Wvalor9 + "','" _
    rem                     + WEnsayo10 + "','" _
    rem                     + Wvalor10 + "','" _
    rem                     + WDate + "'"
    Rem                 Set rstEspecificaciones = db.OpenRecordset("ModificaEspecificaciones " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem
    Rem                         Else
    Rem
    Rem                 XParam = "'" + WProducto + "','" _
    rem                     + WEnsayo1 + "','" _
    rem                     + WValor1 + "','" _
    rem                     + WEnsayo2 + "','" _
    rem                     + Wvalor2 + "','" _
    rem                     + WEnsayo3 + "','" _
    rem                     + WValor3 + "','" _
    rem                     + WEnsayo4 + "','" _
    rem                     + Wvalor4 + "','" _
    rem                     + WEnsayo5 + "','" _
    rem                     + Wvalor5 + "','" _
    rem                     + WEnsayo6 + "','" _
    rem                     + Wvalor6 + "','" _
    rem                     + WEnsayo7 + "','" _
    rem                     + Wvalor7 + "','" _
    rem                     + WEnsayo8 + "','" _
    rem                     + Wvalor8 + "','" _
    rem                     + WEnsayo9 + "','" _
    rem                     + Wvalor9 + "','" _
    rem                     + WEnsayo10 + "','" _
    rem                     + Wvalor10 + "','" _
    rem                     + WDate + "'"
    Rem                 Set rstEspecificaciones = db.OpenRecordset("AltaEspecificaciones " + XParam, dbOpenSnapshot, dbSQLPassThrough)
    Rem             End If
    Rem
    Rem             .MoveNext
    Rem             If .EOF = True Then
    Rem                 Exit Do
    Rem             End If
    Rem         Loop
    Rem End With

    'moviienmtos varios de laboratiorio
        
    coderr = 0
    With rstWMovlab
            .Index = "Clave"
            .MoveFirst
            Do
                
                WCodigo = Str$(!Codigo)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WFechaord = !FechaOrd
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = Str$(!Cantidad)
                WMovi = !Movi
                WTipomov = !Tipomov
                WObservaciones = !Observaciones
                WClave = !Clave
                WDate = ""
                WLote = Str$(!Lote)
                
                spmovlab = "Consultamovlab " + "'" + WClave + "'"
                Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovlab.RecordCount > 0 Then
                    rstMovlab.Close
                    XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WTipomov + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                    spmovlab = "Modificamovlab " + XParam
                    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                            
                    XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WTipomov + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                    spmovlab = "Modificamovlab " + XParam
                    Set rstMovlab = db.OpenRecordset(spmovlab, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    

    Rem PRODUCTOS TERMINADOS
    
    coderr = 0
    With rstWTerminado
            .Index = "Codigo"
            .MoveFirst
            Do
                
                WCodigo = !Codigo
                WDescripcion = !Descripcion
                WLinea = Str$(!Linea)
                WUnidad = !Unidad
                WInicial = Str$(!Inicial)
                WEntradas = Str$(!Entradas)
                WSalidas = Str$(!Salidas)
                WMinimo = Str$(!Minimo)
                WDeposito = !Deposito
                WPedido = Str$(!Pedido)
                WEnvase1 = Str$(!Envase1)
                WEnvase2 = Str$(!Envase2)
                WEnvase3 = Str$(!Envase3)
                WEnvase4 = Str$(!Envase4)
                WEnvase5 = Str$(!Envase5)
                WEnvase6 = Str$(!Envase6)
                WProceso = Str$(!Proceso)
                WCosto = Str$(!Costo)
                WFactor = Str$(!Factor)
                WDate = ""
                WImpreadi = ""
                WClase = ""
                WControla = "0"
                WObservaciones = ""
                
                spTerminado = "ConsultaTerminado " + "'" + WCodigo + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                
                    rstTerminado.Close
                
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
                         + WFechaversion + "','" _
                         + WControla + "','" _
                         + WObservaciones + "'"
                         
                    spTerminado = "ModificaTerminado " + XParam
                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                    
                            Else
                
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
                         + WWDate + "','" _
                         + WImpreadi + "','" _
                         + WClase + "','" _
                         + WIntervencion + "','" _
                         + WNaciones + "','" _
                         + WEmbalaje + "','" _
                         + WVersion + "','" _
                         + WFechaversion + "','" _
                         + WControla + "','" _
                         + WObservaciones + "'"

                    Set rstTerminado = db.OpenRecordset("AltaTerminado " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem Composicion
        
    coderr = 0
    With rstWComposicion
            .Index = "Clave"
            .MoveFirst
            Do
                                        
                WTerminado = !Terminado
                WRenglon = !Renglon
                WTipo = !Tipo
                WArticulo1 = !Articulo1
                WArticulo2 = !Articulo2
                WCantidad = Str$(!Cantidad)
                WClave = !Clave
                WDate = ""
                WCosto1 = "0"
                WCosto2 = "0"
                
                spComposicion = "ConsultaComposicion " + "'" + WClave + "'"
                Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                If rstComposicion.RecordCount > 0 Then
                    rstComposicion.Close
                    XParam = "'" + WClave + "','" + WTerminado + "','" + WRenglon + "','" _
                                     + WTipo + "','" + WArticulo1 + "','" + WArticulo2 + "','" _
                                     + WCantidad + "','" + WDate + "','" + WCosto1 + "','" + WCosto2 + "'"
                    spComposicion = "ModificaComposicion " + XParam
                    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                        Else
                    XParam = "'" + WClave + "','" + WTerminado + "','" + WRenglon + "','" _
                                     + WTipo + "','" + WArticulo1 + "','" + WArticulo2 + "','" _
                                     + WCantidad + "','" + WDate + "','" + WCosto1 + "','" + WCosto2 + "'"
                    spComposicion = "AltaComposicion " + XParam
                    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
   Rem Articulo
   
    coderr = 0
    With rstWArticulo
            .Index = "Codigo"
            .MoveFirst
            Do
            
                WCodigo = !Codigo
                WDescripcion = !Descripcion
                WCosto1 = Str$(!Costo1)
                WCosto2 = Str$(!Costo2)
                WInicial = Str$(!Inicial)
                WEntradas = Str$(!Entradas)
                WSalidas = Str$(!Salidas)
                WMinimo = Str$(!Minimo)
                WLaboratorio = Str$(!Laboratorio)
                WUnidad = !Unidad
                WPedido = Str$(!Pedido)
                WDeposito = !Deposito
                WEnvase = Str$(!Envase)
                WRs = !Rs
                WProveedor = IIf(IsNull(!Proveedor), "0", !Proveedor)
                WDate = ""
                WFlete = ""
                WMoneda = ""
                WControla = "0"
                WFecha = ""
                WOrden = ""
                WDife = "0"
                
                spArticulo = "ConsultaArticulo " + "'" + WCodigo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    rstArticulo.Close
                        
                    XParam = "'" + WCodigo + "','" _
                         + WInicial + "','" _
                         + WEntradas + "','" _
                         + WSalidas + "','" _
                         + WLaboratorio + "','" _
                         + WPedido + "','" _
                         + WDate + "'"
                        
                    spArticulo = "ModificaArticuloMovi " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            
                        Else
                    
                    XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + WCosto1 + "','" _
                         + WCosto2 + "','" _
                         + WInicial + "','" _
                         + WEntradas + "','" _
                         + WSalidas + "','" _
                         + WMinimo + "','" _
                         + WLaboratorio + "','" _
                         + WUnidad + "','" _
                         + WPedido + "','" _
                         + WDeposito + "','" _
                         + WEnvase + "','" _
                         + WRs + "','" _
                         + WFecha + "','" _
                         + WOrden + "','" _
                         + WDife + "','" _
                         + WProveedor + "','" _
                         + WDate + "','" _
                         + WFlete + "','" _
                         + WMoneda + "','" _
                         + WControla + "'"
                                         
                    Set rstArticulo = db.OpenRecordset("AltaArticulo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With

    If Val(WEmpresa) = 2 Then
    
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        coderr = 0
        With rstWArticulo
            .Index = "Codigo"
            .MoveFirst
            Do
            
                WCodigo = !Codigo
                WDescripcion = !Descripcion
                WCosto1 = Str$(!Costo1)
                WCosto2 = Str$(!Costo2)
                WInicial = Str$(!Inicial)
                WEntradas = Str$(!Entradas)
                WSalidas = Str$(!Salidas)
                WMinimo = Str$(!Minimo)
                WLaboratorio = Str$(!Laboratorio)
                WUnidad = !Unidad
                WPedido = Str$(!Pedido)
                WDeposito = !Deposito
                WEnvase = Str$(!Envase)
                WRs = !Rs
                WProveedor = !Proveedor
                WDate = ""
                WFlete = ""
                WMoneda = ""
                WControla = "0"
                WFecha = ""
                WOrden = ""
                WDife = "0"
                
                spArticulo = "ConsultaArticulo " + "'" + WCodigo + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
            
                    WCodigo = rstArticulo!Codigo
                    WDescripcion = rstArticulo!Descripcion
                    WCosto1 = WCosto1
                    WCosto2 = WCosto2
                    WInicial = Str$(rstArticulo!Inicial)
                    WEntradas = Str$(rstArticulo!Entradas)
                    WSalidas = Str$(rstArticulo!Salidas)
                    WMinimo = Str$(rstArticulo!Minimo)
                    WLaboratorio = Str$(rstArticulo!Laboratorio)
                    WUnidad = rstArticulo!Unidad
                    WPedido = Str$(rstArticulo!Pedido)
                    WDeposito = rstArticulo!Deposito
                    WEnvase = Str$(rstArticulo!Envase)
                    WRs = rstArticulo!Rs
                    WFecha = rstArticulo!Fecha
                    WOrden = Str$(rstArticulo!Orden)
                    WDife = Str$(rstArticulo!Dife)
                    WProveedor = rstArticulo!Proveedor
                    WDate = Date$
                    WFlete = Str$(rstArticulo!Flete)
                    WMoneda = rstArticulo!Moneda
                    WControla = Str$(rstArticulo!Controla)
                    
                    rstArticulo.Close
                
                    XParam = "'" + WCodigo + "','" _
                         + WDescripcion + "','" _
                         + WCosto1 + "','" _
                         + WCosto2 + "','" _
                         + WInicial + "','" _
                         + WEntradas + "','" _
                         + WSalidas + "','" _
                         + WMinimo + "','" _
                         + WLaboratorio + "','" _
                         + WUnidad + "','" _
                         + WPedido + "','" _
                         + WDeposito + "','" _
                         + WEnvase + "','" _
                         + WRs + "','" _
                         + WFecha + "','" _
                         + WOrden + "','" _
                         + WDife + "','" _
                         + WProveedor + "','" _
                         + WDate + "','" _
                         + WFlete + "','" _
                         + WMoneda + "','" _
                         + WControla + "'"
                         
                    spArticulo = "ModificaArticulo " + XParam
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
    
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If
    
    Rem ESTADISTICA
        
    coderr = 0
    With rstWEstadistica
            .Index = "Clave"
            .MoveFirst
            Do
                
                WTipo = Str$(!Tipo)
                WNumero = Str$(!Numero)
                WRenglon = Str$(!Renglon)
                WArticulo = !Articulo
                WCantidad = Str$(!Cantidad)
                WPrecio = Str$(!Precio)
                WPrecioUs = Str$(!PrecioUs)
                WImporte = Str$(!Importe)
                WimporteUs = Str$(!ImporteUs)
                WCliente = !Cliente
                WParidad = Str$(!Paridad)
                WVendedor = Str$(!Vendedor)
                WRubro = Str$(!Rubro)
                WLinea = Str$(!Linea)
                WCosto1 = "0"
                WCosto2 = "0"
                WCoeficiente = Str$(!Coeficiente)
                WPedido = Str$(!Pedido)
                WFecha = !Fecha
                WImporte1 = Str$(!Importe1)
                WImporte2 = Str$(!Importe2)
                WImporte3 = Str$(!Importe3)
                WImporte4 = Str$(!Importe4)
                WOrdFecha = !OrdFecha
                WWArticulo = !WArticulo
                WRemito = !Remito
                WClave = !Clave
                WLote1 = Str$(!Lote1)
                WCanti1 = Str$(!Canti1)
                WLote2 = Str$(!Lote2)
                WCanti2 = Str$(!Canti2)
                WLote3 = Str$(!Lote3)
                WCanti3 = Str$(!Canti3)
                WLote4 = Str$(!Lote4)
                WCanti4 = Str$(!Canti4)
                WLote5 = Str$(!Lote5)
                WCanti5 = Str$(!Canti5)
                
                spEstadistica = "ConsultaEstadistica " + "'" + WClave + "'"
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then
                        rstEstadistica.Close
                
                        XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + WRenglon + "','" + WArticulo + "','" _
                                + WCantidad + "','" + WPrecio + "','" _
                                + WPrecioUs + "','" + WImporte + "','" _
                                + WimporteUs + "','" + WCliente + "','" _
                                + WParidad + "','" + WVendedor + "','" _
                                + WRubro + "','" + WLinea + "','" _
                                + WCosto1 + "','" + WCosto2 + "','" _
                                + WCoeficiente + "','" + WPedido + "','" _
                                + WFecha + "','" + WImporte1 + "','" _
                                + WImporte2 + "','" + WImporte3 + "','" _
                                + WImporte4 + "','" + WOrdFecha + "','" _
                                + WWArticulo + "','" + WRemito + "','" _
                                + WDate + "','" + WCanti + "','" _
                                + WImpo + "','" + WImpoUs + "','" _
                                + WMarca + "','" _
                                + WLote1 + "','" + WCanti1 + "','" _
                                + WLote2 + "','" + WCanti2 + "','" _
                                + WLote3 + "','" + WCanti3 + "','" _
                                + WLote4 + "','" + WCanti4 + "','" _
                                + WLote5 + "','" + WCanti5 + "'"
                
                        spEstadistica = "ModificaEstadistica " + XParam
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                            
                                    Else
                                    
                        XParam = "'" + WClave + "','" _
                                + WTipo + "','" + WNumero + "','" _
                                + WRenglon + "','" + WArticulo + "','" _
                                + WCantidad + "','" + WPrecio + "','" _
                                + WPrecioUs + "','" + WImporte + "','" _
                                + WimporteUs + "','" + WCliente + "','" _
                                + WParidad + "','" + WVendedor + "','" _
                                + WRubro + "','" + WLinea + "','" _
                                + WCosto1 + "','" + WCosto2 + "','" _
                                + WCoeficiente + "','" + WPedido + "','" _
                                + WFecha + "','" + WImporte1 + "','" _
                                + WImporte2 + "','" + WImporte3 + "','" _
                                + WImporte4 + "','" + WOrdFecha + "','" _
                                + WWArticulo + "','" + WRemito + "','" _
                                + WDate + "','" + WCanti + "','" _
                                + WImpo + "','" + WImpoUs + "','" _
                                + WMarca + "','" _
                                + WLote1 + "','" + WCanti1 + "','" _
                                + WLote2 + "','" + WCanti2 + "','" _
                                + WLote3 + "','" + WCanti3 + "','" _
                                + WLote4 + "','" + WCanti4 + "','" _
                                + WLote5 + "','" + WCanti5 + "'"
                    
                        spEstadistica = "AltaEstadistica " + XParam
                        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem INFORME DE RECEPCION
        
    coderr = 0
    With rstWInforme
            .Index = "Clave"
            .MoveFirst
            Do
                
                WInforme = Str$(!Informe)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WProveedor = !Proveedor
                WRemito = Str$(!Remito)
                WOrden = Str$(!Orden)
                WArticulo = !Articulo
                WCantidad = Str$(!Cantidad)
                WResta = Str$(!Resta)
                WClave = !Clave
                WFechaord = !FechaOrd
                WDate = ""
                WEnvase = ""
                
                spInforme = "ConsultaInforme " + "'" + WClave + "'"
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    rstInforme.Close
                
                    XParam = "'" + WClave + "','" _
                         + WInforme + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WRemito + "','" _
                         + WProveedor + "','" _
                         + WOrden + "','" _
                         + WArticulo + "','" _
                         + WCantidad + "','" _
                         + WResta + "','" _
                         + WFechaord + "','" _
                         + WDate + "','" _
                         + WEnvase + "'"
                         
                    spInforme = "ModificaInforme " + XParam
                    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                    
                    XParam = "'" + WClave + "','" _
                         + WInforme + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WRemito + "','" _
                         + WProveedor + "','" _
                         + WOrden + "','" _
                         + WArticulo + "','" _
                         + WCantidad + "','" _
                         + WResta + "','" _
                         + WFechaord + "','" _
                         + WDate + "','" _
                         + WEnvase + "'"
                         
                    spInforme = "AltaInforme " + XParam
                    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem LAUDO DE LIBERACION
        
    coderr = 0
    With rstWLaudo
            .Index = "Clave"
            .MoveFirst
            Do
                                    
                WLaudo = Str$(!Laudo)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WOrden = Str$(!Orden)
                WArticulo = !Articulo
                WLiberada = Str$(!Liberada)
                WDevuelta = Str$(!Devuelta)
                WLote = Str$(!Lote)
                WRechazo = Str$(!Rechazo)
                WActualiza = ""
                WMarca = !Marca
                WInforme = Str$(!Informe)
                WClave = !Clave
                WDate = ""
                WSaldo = Str$(!Saldo)
                
                spLaudo = "ConsultaLaudo " + "'" + WClave + "'"
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    rstLaudo.Close
                
                    XParam = "'" + WClave + "','" _
                         + WLaudo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WArticulo + "','" _
                         + WLiberada + "','" _
                         + WDevuelta + "','" _
                         + WOrden + "','" _
                         + WMarca + "','" _
                         + WLote + "','" _
                         + WRechazo + "','" _
                         + WInforme + "','" _
                         + WActualiza + "','" _
                         + WDate + "','" _
                         + WSaldo + "'"
                         
                    spLaudo = "ModificaLaudo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XParam = "'" + WClave + "','" _
                         + WLaudo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WArticulo + "','" _
                         + WLiberada + "','" _
                         + WDevuelta + "','" _
                         + WOrden + "','" _
                         + WMarca + "','" _
                         + WLote + "','" _
                         + WRechazo + "','" _
                         + WInforme + "','" _
                         + WActualiza + "','" _
                         + WDate + "','" _
                         + WSaldo + "'"
                         
                    spLaudo = "AltaLaudo " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
        
    
   Rem MOVIIMENTOS VARIOS
        
    coderr = 0
    With rstWMovvar
            .Index = "Clave"
            .MoveFirst
            Do
                
                WCodigo = Str$(!Codigo)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WFechaord = !FechaOrd
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = Str$(!Cantidad)
                WMovi = !Movi
                WTipomov = !Tipomov
                WObservaciones = !Observaciones
                WClave = !Clave
                WDate = ""
                WLote = Str$(!Lote)
                
                spMovvar = "ConsultaMovvar " + "'" + WClave + "'"
                Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovvar.RecordCount > 0 Then
                    rstMovvar.Close
                
                    XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WTipomov + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                    spMovvar = "ModificaMovvar " + XParam
                    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WFechaord + "','" _
                         + WMovi + "','" _
                         + WTipomov + "','" _
                         + WObservaciones + "','" _
                         + WDate + "','" _
                         + WMarca + "','" _
                         + WLote + "'"
                         
                    spMovvar = "AltaMovvar " + XParam
                    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    
   Rem MOVIIMENTOS VARIOS

    coderr = 0
    With rstWMovguia
            .Index = "Clave"
            .MoveFirst
            Do
                
                WCodigo = Str$(!Codigo)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WFechaord = !FechaOrd
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = Str$(!Cantidad)
                WMovi = !Movi
                WTipomov = !Tipomov
                WObservaciones = !Observaciones
                WClave = !Clave
                WDestino = ""
                WDestino = Str$(!Destino)
                WDate = ""
                WMarca = ""
                WLote = Str$(!Lote)
                WSaldo = Str$(!Saldo)
                WPartida = Str$(!Partida)
                XPartida = Str$(!Partida)
                
                spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                If rstMovguia.RecordCount > 0 Then
                    rstMovguia.Close
                        Else
                        
                    XParam = "'" + WClave + "','" _
                            + WTipomov + "','" _
                            + WCodigo + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WTipo + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WCantidad + "','" _
                            + WFechaord + "','" _
                            + WMovi + "','" _
                            + WObservaciones + "','" _
                            + WDate + "','" _
                            + WMarca + "','" _
                            + WDestino + "','" _
                            + WLote + "','" _
                            + WSaldo + "','" _
                            + WPartida + "'"
                         
                    spMovguia = "AltaMovguia " + XParam
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                End If
                
                If Val(!Destino) = 3 Then
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    Auxi = Str$(!Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(!Codigo)
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(XEmpresa)
                    Call Ceros(WTipomov, 1)
                
                    WClave = WTipomov + Auxi1 + Auxi
                    WCodigo = Str$(!Codigo)
                    WRenglon = Str$(!Renglon)
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WTipo = !Tipo
                    WArticulo = !Articulo
                    WTerminado = !Terminado
                    WCantidad = Str$(!Cantidad)
                    WMovi = "E"
                    WObservaciones = !Observaciones
                    WDestino = "0"
                    WDate = Date$
                    WMarca = ""
                    WLote = XPartida
                    WSaldo = Str$(!Cantidad)
                    WPartida = ""
                
                    spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                            Else
                        Select Case WTipo
                            Case "M"
                                WEntra = "N"
        
                                XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case "T"
                                WEntra = "N"
            
                                XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstHoja.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case Else
                        End Select
                        
                        If WEntra = "S" Then
                            WSaldo = "0"
                        End If
                
                        XParam = "'" + WClave + "','" _
                                    + WTipomov + "','" _
                                    + WCodigo + "','" _
                                    + WRenglon + "','" _
                                    + WFecha + "','" _
                                    + WTipo + "','" _
                                    + WArticulo + "','" _
                                    + WTerminado + "','" _
                                    + WCantidad + "','" _
                                    + WFechaord + "','" _
                                    + WMovi + "','" _
                                    + WObservaciones + "','" _
                                    + WDate + "','" _
                                    + WMarca + "','" _
                                    + WDestino + "','" _
                                    + WLote + "','" _
                                    + WSaldo + "','" _
                                    + WPartida + "'"
                         
                        spMovguia = "AltaMovguia " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case WTipo
                            Case "M"
                                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    WCodigo = WArticulo
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstArticulo!Salidas)
                                                Else
                                        WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstArticulo!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstArticulo.Close
                    
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spArticulo = "ModificaArticuloMovimientos " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                            Case "T"
                                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
        
                                    WCodigo = WTerminado
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstTerminado!Salidas)
                                            Else
                                        WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstTerminado!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstTerminado.Close
                
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                End If
            
                            Case Else
                        End Select
                    End If
                        
                    Select Case Val(XEmpresa)
                        Case 1
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            WEmpresa = "0002"
                            txtOdbc = "Empresa02"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                    
                End If
                
                If Val(!Destino) = 4 Then
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    Auxi = Str$(!Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(!Codigo)
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(XEmpresa)
                    Call Ceros(WTipomov, 1)
                
                    WClave = WTipomov + Auxi1 + Auxi
                    WCodigo = Str$(!Codigo)
                    WRenglon = Str$(!Renglon)
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WTipo = !Tipo
                    WArticulo = !Articulo
                    WTerminado = !Terminado
                    WCantidad = Str$(!Cantidad)
                    WMovi = "E"
                    WObservaciones = !Observaciones
                    WDestino = "0"
                    WDate = Date$
                    WMarca = ""
                    WLote = XPartida
                    WSaldo = Str$(!Cantidad)
                    WPartida = ""
                
                    spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                            Else
                        Select Case WTipo
                            Case "M"
                                WEntra = "N"
        
                                XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case "T"
                                WEntra = "N"
            
                                XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstHoja.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case Else
                        End Select
                        
                        If WEntra = "S" Then
                            WSaldo = "0"
                        End If
                
                        XParam = "'" + WClave + "','" _
                                    + WTipomov + "','" _
                                    + WCodigo + "','" _
                                    + WRenglon + "','" _
                                    + WFecha + "','" _
                                    + WTipo + "','" _
                                    + WArticulo + "','" _
                                    + WTerminado + "','" _
                                    + WCantidad + "','" _
                                    + WFechaord + "','" _
                                    + WMovi + "','" _
                                    + WObservaciones + "','" _
                                    + WDate + "','" _
                                    + WMarca + "','" _
                                    + WDestino + "','" _
                                    + WLote + "','" _
                                    + WSaldo + "','" _
                                    + WPartida + "'"
                         
                        spMovguia = "AltaMovguia " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case WTipo
                            Case "M"
                                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    WCodigo = WArticulo
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstArticulo!Salidas)
                                                Else
                                        WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstArticulo!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstArticulo.Close
                    
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spArticulo = "ModificaArticuloMovimientos " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                            Case "T"
                                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
        
                                    WCodigo = WTerminado
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstTerminado!Salidas)
                                            Else
                                        WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstTerminado!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstTerminado.Close
                
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                End If
            
                            Case Else
                        End Select
                    End If
                        
                    Select Case Val(XEmpresa)
                        Case 1
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            WEmpresa = "0002"
                            txtOdbc = "Empresa02"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                    
                End If
                
                If Val(!Destino) = 5 Then
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    Auxi = Str$(!Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(!Codigo)
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(XEmpresa)
                    Call Ceros(WTipomov, 1)
                
                    WClave = WTipomov + Auxi1 + Auxi
                    WCodigo = Str$(!Codigo)
                    WRenglon = Str$(!Renglon)
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WTipo = !Tipo
                    WArticulo = !Articulo
                    WTerminado = !Terminado
                    WCantidad = Str$(!Cantidad)
                    WMovi = "E"
                    WObservaciones = !Observaciones
                    WDestino = "0"
                    WDate = Date$
                    WMarca = ""
                    WLote = XPartida
                    WSaldo = Str$(!Cantidad)
                    WPartida = ""
                
                    spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                            Else
                        Select Case WTipo
                            Case "M"
                                WEntra = "N"
        
                                XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case "T"
                                WEntra = "N"
            
                                XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstHoja.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case Else
                        End Select
                        
                        If WEntra = "S" Then
                            WSaldo = "0"
                        End If
                
                        XParam = "'" + WClave + "','" _
                                    + WTipomov + "','" _
                                    + WCodigo + "','" _
                                    + WRenglon + "','" _
                                    + WFecha + "','" _
                                    + WTipo + "','" _
                                    + WArticulo + "','" _
                                    + WTerminado + "','" _
                                    + WCantidad + "','" _
                                    + WFechaord + "','" _
                                    + WMovi + "','" _
                                    + WObservaciones + "','" _
                                    + WDate + "','" _
                                    + WMarca + "','" _
                                    + WDestino + "','" _
                                    + WLote + "','" _
                                    + WSaldo + "','" _
                                    + WPartida + "'"
                         
                        spMovguia = "AltaMovguia " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case WTipo
                            Case "M"
                                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    WCodigo = WArticulo
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstArticulo!Salidas)
                                                Else
                                        WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstArticulo!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstArticulo.Close
                    
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spArticulo = "ModificaArticuloMovimientos " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                            Case "T"
                                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
        
                                    WCodigo = WTerminado
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstTerminado!Salidas)
                                            Else
                                        WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstTerminado!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstTerminado.Close
                
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                End If
            
                            Case Else
                        End Select
                    End If
                        
                    Select Case Val(XEmpresa)
                        Case 1
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            WEmpresa = "0002"
                            txtOdbc = "Empresa02"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                    
                End If
                
                If Val(!Destino) = 6 Then
                
                    XEmpresa = WEmpresa
    
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                    Auxi = Str$(!Renglon)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(!Codigo)
                    Call Ceros(Auxi1, 6)
                
                    WTipomov = Str$(XEmpresa)
                    Call Ceros(WTipomov, 1)
                
                    WClave = WTipomov + Auxi1 + Auxi
                    WCodigo = Str$(!Codigo)
                    WRenglon = Str$(!Renglon)
                    WFecha = !Fecha
                    WFechaord = !FechaOrd
                    WTipo = !Tipo
                    WArticulo = !Articulo
                    WTerminado = !Terminado
                    WCantidad = Str$(!Cantidad)
                    WMovi = "E"
                    WObservaciones = !Observaciones
                    WDestino = "0"
                    WDate = Date$
                    WMarca = ""
                    WLote = XPartida
                    WSaldo = Str$(!Cantidad)
                    WPartida = ""
                
                    spMovguia = "ConsultaMovguia " + "'" + WClave + "'"
                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                    If rstMovguia.RecordCount > 0 Then
                        rstMovguia.Close
                            Else
                        Select Case WTipo
                            Case "M"
                                WEntra = "N"
        
                                XParam = "'" + WLote + "','" _
                                         + WArticulo + "'"
                                spLaudo = "ListaLaudoArticulo " + XParam
                                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstLaudo.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstLaudo!Clave
                                    WSaldo = Str$(rstLaudo!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstLaudo.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spLaudo = "ModificaLaudoSaldo " + XParam
                                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WArticulo + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case "T"
                                WEntra = "N"
            
                                XParam = "'" + WLote + "','" _
                                        + WTerminado + "'"
                                spHoja = "ListaHojaProducto " + XParam
                                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                If rstHoja.RecordCount > 0 Then
                                    WEntra = "S"
                                    XClave = rstHoja!Clave
                                    WSaldo = Str$(rstHoja!Saldo + Val(WCantidad))
                                    WDate = Date$
                                    rstHoja.Close
                            
                                    XParam = "'" + XClave + "','" _
                                            + WDate + "','" _
                                            + WSaldo + "'"
                                    spHoja = "ModificaHojaSaldo " + XParam
                                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                                If WEntra = "N" Then
                                    XParam = "'" + WTerminado + "','" _
                                            + WLote + "'"
                                    spMovguia = "ListaMovguiaLote1 " + XParam
                                    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    If rstMovguia.RecordCount > 0 Then
                                        WEntra = "S"
                                        XClave = rstMovguia!Clave
                                        WSaldo = Str$(rstMovguia!Saldo + Val(WCantidad))
                                        WDate = Date$
                                        rstMovguia.Close
                            
                                        XParam = "'" + XClave + "','" _
                                                + WDate + "','" _
                                                + WSaldo + "'"
                                        spMovguia = "ModificaMovguiaSaldo " + XParam
                                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                                    End If
                                End If
                            Case Else
                        End Select
                        
                        If WEntra = "S" Then
                            WSaldo = "0"
                        End If
                
                        XParam = "'" + WClave + "','" _
                                    + WTipomov + "','" _
                                    + WCodigo + "','" _
                                    + WRenglon + "','" _
                                    + WFecha + "','" _
                                    + WTipo + "','" _
                                    + WArticulo + "','" _
                                    + WTerminado + "','" _
                                    + WCantidad + "','" _
                                    + WFechaord + "','" _
                                    + WMovi + "','" _
                                    + WObservaciones + "','" _
                                    + WDate + "','" _
                                    + WMarca + "','" _
                                    + WDestino + "','" _
                                    + WLote + "','" _
                                    + WSaldo + "','" _
                                    + WPartida + "'"
                         
                        spMovguia = "AltaMovguia " + XParam
                        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
                        
                        Select Case WTipo
                            Case "M"
                                spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
                                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                If rstArticulo.RecordCount > 0 Then
                                    WCodigo = WArticulo
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstArticulo!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstArticulo!Salidas)
                                                Else
                                        WSalidas = Str$(rstArticulo!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstArticulo!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstArticulo.Close
                    
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spArticulo = "ModificaArticuloMovimientos " + XParam
                                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                                End If
                
                            Case "T"
                                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                If rstTerminado.RecordCount > 0 Then
        
                                    WCodigo = WTerminado
                                    If WMovi = "E" Then
                                        WEntradas = Str$(rstTerminado!Entradas + Val(WCantidad))
                                        WSalidas = Str$(rstTerminado!Salidas)
                                            Else
                                        WSalidas = Str$(rstTerminado!Salidas + Val(WCantidad))
                                        WEntradas = Str$(rstTerminado!Entradas)
                                    End If
                                    WDate = Date$
                                    
                                    rstTerminado.Close
                
                                    XParam = "'" + WCodigo + "','" _
                                            + WEntradas + "','" _
                                            + WSalidas + "','" _
                                            + WDate + "'"
                                           
                                    spTerminado = "ModificaTerminadoMovimientos " + XParam
                                    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                                End If
            
                            Case Else
                        End Select
                    End If
                        
                    Select Case Val(XEmpresa)
                        Case 1
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case 2
                            WEmpresa = "0002"
                            txtOdbc = "Empresa02"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                    End Select
                    
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
   Rem HOJA DE PRODUCCION
        
    coderr = 0
    With rstWHoja
            .Index = "Clave"
            .MoveFirst
            Do
                
                WHoja = Str$(!Hoja)
                WRenglon = Str$(!Renglon)
                WFecha = !Fecha
                WProducto = !Producto
                WTeorico = Str$(!Teorico)
                WReal = Str$(!Real)
                WFechaing = !Fechaing
                WFechaingord = Right$(!Fechaing, 4) + Mid$(!Fechaing, 4, 2) + Left$(!Fechaing, 2)
                WTipo = !Tipo
                WArticulo = !Articulo
                WTerminado = !Terminado
                WCantidad = Str$(!Cantidad)
                WLote = Str$(!Lote)
                WClave = !Clave
                WMarca = ""
                WSaldo = Str$(!Saldo)
                WLote1 = Str$(!Lote1)
                WCanti1 = Str$(!Canti1)
                WLote2 = Str$(!Lote2)
                WCanti2 = Str$(!Canti2)
                WLote3 = Str$(!Lote3)
                WCanti3 = Str$(!Canti3)
                
                spHoja = "ConsultaHoja " + "'" + WClave + "'"
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                        rstHoja.Close
                
                        XParam = "'" + WClave + "','" _
                            + WHoja + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WProducto + "','" _
                            + WCantidad + "','" _
                            + WTipo + "','" _
                            + WLote + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WTeorico + "','" _
                            + WReal + "','" _
                            + WFechaing + "','" _
                            + WFechaingord + "','" _
                            + WDate + "','" _
                            + WImporte + "','" _
                            + WMarca + "','" _
                            + WSaldo + "','" _
                            + WLote1 + "','" _
                            + WCanti1 + "','" _
                            + WLote2 + "','" _
                            + WCanti2 + "','" _
                            + WLote3 + "','" _
                            + WCanti3 + "'"
                                           
                        spHoja = "ModificaHoja " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                        
                            Else
                            
                        XParam = "'" + WClave + "','" _
                            + WHoja + "','" _
                            + WRenglon + "','" _
                            + WFecha + "','" _
                            + WProducto + "','" _
                            + WCantidad + "','" _
                            + WTipo + "','" _
                            + WLote + "','" _
                            + WArticulo + "','" _
                            + WTerminado + "','" _
                            + WTeorico + "','" _
                            + WReal + "','" _
                            + WFechaing + "','" _
                            + WFechaingord + "','" _
                            + WDate + "','" _
                            + WImporte + "','" _
                            + WMarca + "','" _
                            + WSaldo + "','" _
                            + WLote1 + "','" _
                            + WCanti1 + "','" _
                            + WLote2 + "','" _
                            + WCanti2 + "','" _
                            + WLote3 + "','" _
                            + WCanti3 + "'"
                                           
                        spHoja = "AltaHoja " + XParam
                        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                    End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
    If Val(WEmpresa) = 2 Then
    
        WEmpresa = "0004"
        txtOdbc = "Empresa04"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        coderr = 0
        With rstWCtaCte
            .Index = "Clave"
            .MoveFirst
            Do
                
                WTipo = !Tipo
                WNumero = Str$(!Numero)
                WRenglon = !Renglon
                WCliente = !Cliente
                WFecha = !Fecha
                WEstado = !Estado
                Wvencimiento = !Vencimiento
                WVencimiento1 = !Vencimiento1
                WTotal = Str$(!Total)
                WTotalUs = Str$(!TotalUs)
                WSaldo = Str$(!Saldo)
                WSaldoUs = Str$(!Saldous)
                WOrdFecha = !OrdFecha
                WOrdVencimiento = !OrdVencimiento
                WOrdVencimiento1 = !OrdVencimiento1
                WImpre = !Impre
                WNeto = Str$(!Neto)
                WIva1 = Str$(!Iva1)
                WIva2 = Str$(!Iva2)
                WPedido = !Pedido
                WRemito = !Remito
                WOrden = !Orden
                WParidad = Str$(!Paridad)
                WProvincia = !Provincia
                WVendedor = Str$(!Vendedor)
                WRubro = Str$(!Rubro)
                WComprobante = !Comprobante
                WAceptada = !Aceptada
                WCosto = Str$(!Costo)
                WImporte1 = Str$(!Importe1)
                WImporte2 = Str$(!Importe2)
                WImporte3 = Str$(!Importe3)
                WImporte4 = Str$(!Importe4)
                WImporte5 = Str$(!Importe5)
                WImporte6 = Str$(!Importe6)
                WImporte7 = Str$(!Importe7)
                WClave = !Clave
                
                spCtacte = "ConsultaCtaCte " + "'" + WClave + "'"
                Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                If rstCtacte.RecordCount > 0 Then
                        rstCtacte.Close
                
                        XParam = "'" + WClave + "','" _
                            + WTipo + "','" + WNumero + "','" _
                            + WRenglon + "','" + WCliente + "','" _
                            + WFecha + "','" + WEstado + "','" _
                            + Wvencimiento + "','" + WVencimiento1 + "','" _
                            + WTotal + "','" + WTotalUs + "','" _
                            + WSaldo + "','" + WSaldoUs + "','" _
                            + WOrdFecha + "','" + WOrdVencimiento + "','" _
                            + WOrdVencimiento1 + "','" + WImpre + "','" _
                            + WEmpresa + "','" _
                            + WNeto + "','" + WIva1 + "','" _
                            + WIva2 + "','" + WPedido + "','" _
                            + WRemito + "','" + WOrden + "','" _
                            + WParidad + "','" + WProvincia + "','" _
                            + WVendedor + "','" + WRubro + "','" _
                            + WComprobante + "','" + WAceptada + "','" _
                            + WCosto + "','" _
                            + WImporte1 + "','" _
                            + WImporte2 + "','" _
                            + WImporte3 + "','" _
                            + WImporte4 + "','" _
                            + WImporte5 + "','" _
                            + WImporte6 + "','" _
                            + WImporte7 + "','" _
                            + WDate + "'"
                        
                        spCtacte = "ModificaCtacte " + XParam
                        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                        
                                Else
                                
                        XParam = "'" + WClave + "','" _
                            + WTipo + "','" + WNumero + "','" _
                            + WRenglon + "','" + WCliente + "','" _
                            + WFecha + "','" + WEstado + "','" _
                            + Wvencimiento + "','" + WVencimiento1 + "','" _
                            + WTotal + "','" + WTotalUs + "','" _
                            + WSaldo + "','" + WSaldoUs + "','" _
                            + WOrdFecha + "','" + WOrdVencimiento + "','" _
                            + WOrdVencimiento1 + "','" + WImpre + "','" _
                            + WEmpresa + "','" _
                            + WNeto + "','" + WIva1 + "','" _
                            + WIva2 + "','" + WPedido + "','" _
                            + WRemito + "','" + WOrden + "','" _
                            + WParidad + "','" + WProvincia + "','" _
                            + WVendedor + "','" + WRubro + "','" _
                            + WComprobante + "','" + WAceptada + "','" _
                            + WCosto + "','" _
                            + WImporte1 + "','" _
                            + WImporte2 + "','" _
                            + WImporte3 + "','" _
                            + WImporte4 + "','" _
                            + WImporte5 + "','" _
                            + WImporte6 + "','" _
                            + WImporte7 + "','" _
                            + WDate + "'"
                        
                        spCtacte = "AltaCtacte " + XParam
                        Set rstCtacte = db.OpenRecordset(spCtacte, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        
        WEmpresa = "0002"
        txtOdbc = "Empresa02"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    End If
    
    If Val(WEmpresa) = 1 Then
        Call Alta_Proveedor
    End If
    
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub


Private Sub Proceso1()

    On Error GoTo Error
    
    OPEN_FILE_WOrden
    Rem OPEN_FILE_WCotiza
    
    
   Rem cotizaciones
        
    Rem coderr = 0
    Rem With rstWCotiza
    Rem         .Index = "Clave"
    Rem          .MoveFirst
    Rem          Do
    Rem
    Rem              WCotiza = !Cotiza
    Rem              WRenglon = !Renglon
    Rem              WFecha = !Fecha
    Rem              WProveedor = !Proveedor
    Rem              WArticulo = !Articulo
    Rem              WPrecio = !Precio
    Rem              WFechaord = !Fechaord
    Rem              WCondicion = !Condicion
    Rem              WObservaciones = !Observaciones
    Rem              WClave = !Clave
    Rem
    Rem              With rstCotiza
    Rem                      .Index = "Clave"
    Rem                      .Seek "=", WClave
    Rem                      If .NoMatch Then
    Rem                          .AddNew
    Rem                          !Cotiza = WCotiza
    Rem                          !Renglon = WRenglon
    Rem                          !Fecha = WFecha
    Rem                          !Proveedor = WProveedor
    Rem                          !Articulo = WArticulo
    Rem                          !Precio = WPrecio
    Rem                          !Fechaord = WFechaord
    Rem                          !Condicion = WCondicion
    Rem                          !Observaciones = WObservaciones
    Rem                          !Clave = WClave
    Rem                          .Update
    Rem                          .Bookmark = .LastModified
    Rem                              Else
    Rem                          .Edit
    Rem                          !Cotiza = WCotiza
    Rem                          !Renglon = WRenglon
    Rem                          !Fecha = WFecha
    Rem                          !Proveedor = WProveedor
    Rem                          !Articulo = WArticulo
    Rem                          !Precio = WPrecio
    Rem                          !Fechaord = WFechaord
    Rem                          !Condicion = WCondicion
    Rem                          !Observaciones = WObservaciones
    Rem                          !Clave = WClave
    Rem                          .Update
    Rem                          .Bookmark = .LastModified
    Rem                      End If
    Rem              End With
    Rem
    Rem              .MoveNext
    Rem              If .EOF = True Then
    Rem                  Exit Do
    Rem              End If
    Rem          Loop
    Rem End With
    
   Rem ORDENES DE COMPRA
        
    coderr = 0
    With rstWOrden
             .Index = "Clave"
             .MoveFirst
             Do
    
                 WOrden = Str$(!Orden)
                 WRenglon = Str$(!Renglon)
                 WFecha = !Fecha
                 WFechaord = !FechaOrd
                 WProveedor = !Proveedor
                 WArticulo = !Articulo
                 WCantidad = Str$(!Cantidad)
                 WPrecio = Str$(!Precio)
                 WFecha1 = !Fecha1
                 WFecha2 = !Fecha2
                 WCondicion = !Condicion
                 WRecibida = Str$(!Recibida)
                 WClave = !Clave
                 WSaldo = ""
                 WLiberada = ""
                 WDevuelta = ""
                 WFechaentrega = ""
                 WDate = ""
                 
                 spOrden = "ConsultaOrden " + "'" + WClave + "'"
                 Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                 If rstOrden.RecordCount > 0 Then
                        rstOrden.Close
                 
                        XParam = "'" + WClave + "','" _
                                + WOrden + "','" _
                                + WRenglon + "','" _
                                + WFecha + "','" _
                                + WProveedor + "','" _
                                + WArticulo + "','" _
                                + WCantidad + "','" _
                                + WPrecio + "','" _
                                + WFecha1 + "','" _
                                + WFecha2 + "','" _
                                + WCondicion + "','" _
                                + WRecibida + "','" _
                                + WSaldo + "','" _
                                + WFechaord + "','" _
                                + WLiberada + "','" _
                                + WDevuelta + "','" _
                                + WFechaentrega + "','" _
                                + WDate + "'"
                         
                        Rem spOrden = "ModificaOrden " + XParam
                        Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        
                                Else
                                
                        XParam = "'" + WClave + "','" _
                                + WOrden + "','" _
                                + WRenglon + "','" _
                                + WFecha + "','" _
                                + WProveedor + "','" _
                                + WArticulo + "','" _
                                + WCantidad + "','" _
                                + WPrecio + "','" _
                                + WFecha1 + "','" _
                                + WFecha2 + "','" _
                                + WCondicion + "','" _
                                + WRecibida + "','" _
                                + WSaldo + "','" _
                                + WFechaord + "','" _
                                + WLiberada + "','" _
                                + WDevuelta + "','" _
                                + WFechaentrega + "','" _
                                + WDate + "'"
                         
                        spOrden = "AltaOrden " + XParam
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                .MoveNext
                If .EOF = True Then
                     Exit Do
                End If
            Loop
    End With
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub


Private Sub Form_Load()
    Call Acepta_Click
End Sub


Private Sub Alta_Proveedor()

    
   Rem Proveedor
        
    coderr = 0
    With rstWProveedor
            .Index = "Proveedor"
            .MoveFirst
            Do
            
                WProveedor = !Proveedor
                WNombre = !Nombre
                WDireccion = !Direccion
                WLocalidad = !Localidad
                WPostal = !Postal
                WCuit = !Cuit
                WTelefono = !Telefono
                WEMail = !EMail
                WObservaciones = !Observaciones
                WDias = Str$(!Dias)
                WCuenta = Str$(!Cuenta)
                WIva = !Iva
                WTipo = !Tipo
                WProvincia = !Provincia
                WNombreCheque = ""
                XEmp = "1"
                WImporte1 = ""
                WImporte2 = ""
                WImporte3 = ""
                WImporte4 = ""
                WImporte5 = ""
                WImporte6 = ""
                
                XEmpresa = WEmpresa
                
                For Xx = 1 To 6
                
                    Select Case Xx
                        Case 1
                            WEmpresa = "0001"
                            txtOdbc = "Empresa01"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case 2
                            WEmpresa = "0002"
                            txtOdbc = "Empresa02"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case 3
                            WEmpresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case 4
                            WEmpresa = "0004"
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case 5
                            WEmpresa = "0005"
                            txtOdbc = "Empresa05"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case 6
                            WEmpresa = "0006"
                            txtOdbc = "Empresa06"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                            
                        Case Else
                    End Select
                            
                    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
                    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                    If RstProveedor.RecordCount > 0 Then
                        RstProveedor.Close
                        XParam = "'" + WProveedor + "','" + WNombre + "','" + WDireccion + "','" _
                                + WLocalidad + "','" + WProvincia + "','" + WPostal + "','" + WCuit + "','" _
                                + WTelefono + "','" + WEMail + "','" + WObservaciones + "','" _
                                + WTipo + "','" + WIva + "','" _
                                + WDias + "','" + XEmp + "','" + WCuenta + "','" _
                                + WImporte1 + "','" + WImporte2 + "','" _
                                + WImporte2 + "','" + WImporte4 + "','" _
                                + WImporte3 + "','" + WImporte6 + "','" _
                                + WNombreCheque + "'"
                        Set RstProveedor = db.OpenRecordset("ModificaProveedor " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                
                            Else
        
                        XParam = "'" + WProveedor + "','" + WNombre + "','" + WDireccion + "','" _
                                + WLocalidad + "','" + WProvincia + "','" + WPostal + "','" + WCuit + "','" _
                                + WTelefono + "','" + WEMail + "','" + WObservaciones + "','" _
                                + WTipo + "','" + WIva + "','" _
                                + WDias + "','" + XEmp + "','" + WCuenta + "','" _
                                + WImporte1 + "','" + WImporte2 + "','" _
                                + WImporte2 + "','" + WImporte4 + "','" _
                                + WImporte3 + "','" + WImporte6 + "','" _
                                + WNombreCheque + "'"
                        Set RstProveedor = db.OpenRecordset("AltaProveedor " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                Next Xx
                
                Select Case Val(XEmpresa)
                    Case 1
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        WEmpresa = "0002"
                        txtOdbc = "Empresa02"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 3
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 4
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 5
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 6
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
    End With
    
End Sub


