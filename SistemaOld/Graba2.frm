VERSION 5.00
Begin VB.Form PrgGraba2 
   Caption         =   "Lectura de Prestamos entre Plantas"
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
Attribute VB_Name = "PrgGraba2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Producto As String
Private Costo As Double
Dim XParam As String
Dim WClave As String
Dim WCodigo As String
Dim WRenglon As String
Dim WFecha As String
Dim OrdFecha As String
Dim WObservaciones As String
Dim WTipo As String
Dim WArticulo As String
Dim WTerminado As String
Dim WCantidad As String
Dim WCosto As String
Dim WDestino As String
Dim WTermino As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Private WVector(10000, 10) As String
Private Auxiliar(100, 7) As String

Private Sub Acepta_Click()

    Call Proceso
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgGraba2.Hide
    Unload Me
    End
End Sub

Private Sub Proceso()

    Rem On Error GoTo Error
    
    spMovguia = "ListaMovguiaTotal"
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
    If rstMovguia.RecordCount > 0 Then
        With rstMovguia
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstMovguia!FechaOrd >= "20020101" Then
                        If rstMovguia!Codigo > 90000 Then
                            If rstMovguia!Tipomov = 0 Then
                                Renglon = Renglon + 1
                                WVector(Renglon, 1) = rstMovguia!Tipo
                                WVector(Renglon, 2) = rstMovguia!Articulo
                                WVector(Renglon, 3) = rstMovguia!Terminado
                                WVector(Renglon, 4) = rstMovguia!Cantidad
                                WVector(Renglon, 5) = rstMovguia!destino
                                WVector(Renglon, 6) = rstMovguia!Fecha
                                WVector(Renglon, 7) = rstMovguia!FechaOrd
                                WVector(Renglon, 8) = rstMovguia!Codigo
                            End If
                        End If
                    End If
                    .MoveNext
                            Else
                    Exit Do
                End If
            Loop
        End With
        rstMovguia.Close
    End If
    
    For XCicla = 1 To Renglon
    
        XTipo = WVector(XCicla, 1)
        XArticulo = WVector(XCicla, 2)
        XTerminado = WVector(XCicla, 3)
        XCantidad = WVector(XCicla, 4)
        XDestino = WVector(XCicla, 5)
        XFecha = WVector(XCicla, 6)
        XFechaord = WVector(XCicla, 7)
        XCodigo = WVector(XCicla, 8)
        
        Rem spPrestamo = "ListaPrestamoTotal "
        Rem Set rstPrestamo = db.OpenRecordset(spPrestamo, dbOpenSnapshot, dbSQLPassThrough)
        Rem If rstPrestamo.RecordCount > 0 Then
        Rem     With rstPrestamo
        Rem         .MoveLast
        Rem         XCodigo = rstPrestamo!Codigo + 1
        Rem     End With
        Rem     rstPrestamo.Close
        Rem         Else
        Rem     XCodigo = "1"
        Rem End If

        WCodigo = XCodigo
        WRenglon = "1"
        WFecha = XFecha
        WOrdFecha = XFechaord
        WObservaciones = ""
        WTipo = XTipo
        WArticulo = XArticulo
        WTerminado = XTerminado
        WCantidad = Str$(XCantidad)
        WCosto = ""
        WDestino = ""
                
        Call Ceros(WCodigo, 6)
        Call Ceros(WRenglon, 2)
                
        WClave = WCodigo + WRenglon
        
        XEmpresa = WEmpresa
        Select Case Val(XEmpresa)
            Case 1, 3, 5, 6, 7
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        If WTipo = "M" Then
            
            spArticulo = "ConsultaArticulo" + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCosto = Str$(rstArticulo!Costo1)
                rstArticulo.Close
            End If
            
                Else
                
            Call Calcula_Costo(WTerminado, Costo)
            WCosto = Str$(Costo)
        
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
                Case 7
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 8
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 9
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
        End Select
                
        XParam = "'" + WClave + "','" _
                         + WCodigo + "','" _
                         + WRenglon + "','" _
                         + WFecha + "','" _
                         + WOrdFecha + "','" _
                         + WObservaciones + "','" _
                         + WTipo + "','" _
                         + WArticulo + "','" _
                         + WTerminado + "','" _
                         + WCantidad + "','" _
                         + WCosto + "','" _
                         + WDestino + "'"
                                         
        Set rstPrestamo = db.OpenRecordset("AltaPrestamo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
    Next XCicla
                
    
    Exit Sub
    
Error:
     coderr = Err
     Resume Next
     
End Sub

Private Sub Calcula_Costo(Producto As String, Costo As Double)

    Dim Vector(100, 2) As String
    Erase Auxiliar
    Renglon = 0
    
    Vector(1, 1) = Producto
    Vector(1, 2) = "1"
    Costo = 0
    Lugar = 1
    Cicla = 0
    
    Do
        Cicla = Cicla + 1
        If Vector(Cicla, 1) <> "" Then
    
            spComposicion = "ConsultaComposicionProducto " + "'" + Vector(Cicla, 1) + "'"
            Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
            
            If rstComposicion.RecordCount > 0 Then
            With rstComposicion
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        Tipo = rstComposicion!Tipo
                        Articulo1 = rstComposicion!Articulo1
                        Articulo2 = rstComposicion!Articulo2
                        Cantidad = rstComposicion!Cantidad
                        
                        Select Case Tipo
                            Case "T"
                                If Producto <> Articulo2 Then
                                    Lugar = Lugar + 1
                                    Vector(Lugar, 1) = Articulo2
                                    Vector(Lugar, 2) = Str$(Cantidad * Val(Vector(Cicla, 2)))
                                End If
                            Case "M"
                                Renglon = Renglon + 1
                                Auxiliar(Renglon, 1) = Articulo1
                                Auxiliar(Renglon, 2) = Cantidad
                                Auxiliar(Renglon, 3) = Vector(Cicla, 2)
                            Case Else
                        End Select
                        
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstComposicion.Close
            End If
            
                Else
                
            Exit Do
            
        End If
        
    Loop
                    
    For Da = 1 To Renglon
        Articulo = Auxiliar(Da, 1)
        Cantidad = Auxiliar(Da, 2)
        XVector = Auxiliar(Da, 3)
        
        spArticulo = "ConsultaArticulo " + "'" + Articulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WCosto = (Cantidad * rstArticulo!Costo2 * Val(XVector))
            Costo = Costo + (Cantidad * rstArticulo!Costo2 * Val(XVector))
            rstArticulo.Close
        End If
    Next Da
    
End Sub



