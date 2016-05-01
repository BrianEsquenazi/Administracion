VERSION 5.00
Begin VB.Form PrgCierreStk 
   AutoRedraw      =   -1  'True
   Caption         =   "Cierre de Stock"
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
Attribute VB_Name = "PrgCierreStk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Private Uno As String
Private Dos As String
Private Tres As String
Private Auxi As String
Private Auxi1 As String
Private Auxi2 As String

Private Sub Cancelar_Click()

    PrgCierre.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    WFecha = "01/01/1990"
    
    spHoja = "BorrarHojaFecha " + "'" + WFecha + "'"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenDynaset, dbSQLPassThrough)

    spLaudo = "BorrarLaudoFecha " + "'" + WFecha + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenDynaset, dbSQLPassThrough)

    spLaudo = "ModificaLaudoMarca"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    spHoja = "ModificaHojaMarca"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovvar = "ModificaMovvarMarca"
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovguia = "ModificaMovguiaMarca"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovlab = "ModificaMovlabMarca"
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    
    spEstadistica = "ModificaEstadisticaMarca"
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    
    spArticulo = "ModificaArticuloInicial0"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    spTerminado = "ModificaTerminadoInicial0"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
Stop
    With rstInveMp
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Uno = !X1
                Dos = !X2
                W3 = IIf(IsNull(!X3), "100", !X3)

                If W3 = 0 Then
                    Tres = "100"
                        Else
                    Tres = W3
                End If
                Lote = !X4
                Cantidad = IIf(IsNull(!X5), "0", !X5)
                
                Call Ceros(Dos, 3)
                Call Ceros(Tres, 3)
                Articulo = Uno + "-" + Dos + "-" + Tres
                
                XParam = "'" + Lote + "','" _
                            + Articulo + "'"
                spLaudo = "ListaLaudoArticulo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    WClave = rstLaudo!Clave
                    WSaldo = Str$(rstLaudo!Saldo + Cantidad)
                    WLiberada = Str$(rstLaudo!Liberada + Cantidad)
                    WDate = Date$
                    rstLaudo.Close
                            
                    XParam = "'" + WClave + "','" _
                            + WDate + "','" _
                            + WSaldo + "','" _
                            + WLiberada + "'"
                    spLaudo = "ModificaLaudoSaldo1 " + XParam
                    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                        
                    WLaudo = Str$(Lote)
                    WRenglon = "1"
                    WFecha = "01/01/1990"
                    WOrden = "0"
                    WArticulo = Articulo
                    WLiberada = Str$(Cantidad)
                    WDevuelta = "0"
                    WLote = Str$(Lote)
                    WRechazo = ""
                    WActualiza = "N"
                    WMarca = "X"
                    WInforme = "0"
                    WSaldo = Str$(Cantidad)
            
                    Auxi1 = Str$(WLaudo)
                    Call Ceros(Auxi1, 6)
                    Auxi2 = Str$(WRenglon)
                    Call Ceros(Auxi2, 2)
            
                    WClave = Auxi1 + Auxi2
                    WDate = Date$
        
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
                
                        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                        
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    With rstInvePT
        .Index = "Clave"
        .MoveFirst
        If .NoMatch = False Then
            Do
            
                Uno = !X1
                
                If Uno = "PT" Then
                
                Dos = !X2
                W3 = IIf(IsNull(!X3), "100", !X3)

                If W3 = 0 Then
                    Tres = "100"
                        Else
                    Tres = W3
                End If
                Lote = !X4
                Cantidad = IIf(IsNull(!X5), "0", !X5)
                
                Call Ceros(Dos, 5)
                Call Ceros(Tres, 3)
                Terminado = Uno + "-" + Dos + "-" + Tres
                
                XParam = "'" + Lote + "','" _
                            + Terminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WClave = rstHoja!Clave
                    WSaldo = Str$(rstHoja!Saldo + Cantidad)
                    WReal = Str$(rstHoja!Real + Cantidad)
                    WDate = Date$
                    rstHoja.Close
                            
                    XParam = "'" + WClave + "','" _
                            + WDate + "','" _
                            + WSaldo + "','" _
                            + WReal + "'"
                    spHoja = "ModificaHojaSaldo1 " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                        Else
                    
                    WHoja = Str$(Lote)
                    WRenglon = "1"
                    WFecha = "01/01/1990"
                    WProducto = Terminado
                    WTeorico = "0"
                    WReal = Str$(Cantidad)
                    WFechaing = "00/00/0000"
                    WFechaingord = "00000000"
                    WTipo = "M"
                    WArticulo = "  -   -   "
                    WTerminado = "  -     -   "
                    WCantidad = "0"
                    WLote = ""
                    WDate = Date$
                    WImporte = ""
                    WMarca = "X"
                    WSaldo = Str$(Cantidad)
                    WLote1 = "0"
                    WLote2 = "0"
                    Wlote3 = "0"
                    WCanti1 = "0"
                    WCanti2 = "0"
                    WCanti3 = "0"
                    WCosto1 = "0"
                    WCosto2 = "0"
                    WCosto3 = "0"
                    
                    Auxi = Str$(1)
                    Call Ceros(Auxi, 2)
                        
                    Auxi1 = Str$(Lote)
                    Call Ceros(Auxi1, 6)
                    
                    WClave = Auxi1 + Auxi
                    
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
                            + WLote1 + "','" + WCanti1 + "','" _
                            + WLote2 + "','" + WCanti2 + "','" _
                            + Wlote3 + "','" + Wlote3 + "','" _
                            + WCosto1 + "','" _
                            + WCosto2 + "','" _
                            + WCosto3 + "'"
                                           
                    spHoja = "AltaHoja " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
                
                End If
                
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    
    Call Cancelar_Click

End Sub


Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgCierre.Caption = "Cierre :  " + !Nombre
        End If
    End With

End Sub
