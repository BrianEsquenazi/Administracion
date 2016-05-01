VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgVerilot2Auto 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Saldos de Lotes de Productos Terminados"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlotemat.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   6000
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "verilot2auto.frx":0000
      Left            =   120
      List            =   "verilot2auto.frx":0007
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   5880
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgVerilot2Auto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstGuia As Recordset
Dim spGuia As String
Dim XParam As String
Dim Vector(10000, 5) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double
Dim Empe(12, 10) As String

Private Sub Acepta_Click()

    Empe(1, 1) = "0001"
    Empe(1, 2) = "Empresa01"
    Empe(2, 1) = "0002"
    Empe(2, 2) = "Empresa02"
    Empe(3, 1) = "0003"
    Empe(3, 2) = "Empresa03"
    Empe(4, 1) = "0004"
    Empe(4, 2) = "Empresa04"
    Empe(5, 1) = "0005"
    Empe(5, 2) = "Empresa05"
    Empe(6, 1) = "0006"
    Empe(6, 2) = "Empresa06"
    Empe(7, 1) = "0007"
    Empe(7, 2) = "Empresa07"
    Empe(8, 1) = "0008"
    Empe(8, 2) = "Empresa08"
    Empe(9, 1) = "0009"
    Empe(9, 2) = "Empresa09"
    Empe(10, 1) = "0010"
    Empe(10, 2) = "Empresa10"
    Empe(11, 1) = "0011"
    Empe(11, 2) = "Empresa11"
    
    For A = WDesdeEmpresa To WHastaEmpresa
     
    
    WEmpresa = Empe(A, 1)
    txtOdbc = Empe(A, 2)
    
  
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)


    Erase Vector
    Renglon = 0
    
    Sql1 = "Select Hoja,Marca,Renglon,Producto,Saldo,Real"
    Sql2 = " FROM Hoja"
    Sql3 = " Where Hoja.Marca <> " + "'" + "X" + "'"
    Sql4 = " and Hoja.Renglon = 1 "
    spHoja = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If rstHoja!Marca <> "X" And rstHoja!Renglon = 1 And Left$(rstHoja!producto, 2) = "PT" Then
                            If Left$(rstHoja!producto, 2) = "PT" Or Left$(rstHoja!producto, 2) = "SE" Or Left$(rstHoja!producto, 2) = "YQ" Or Left$(rstHoja!producto, 2) = "YF" Then
                                Renglon = Renglon + 1
                                Vector(Renglon, 1) = rstHoja!Hoja
                                Vector(Renglon, 2) = Str$(rstHoja!Saldo)
                                Vector(Renglon, 3) = rstHoja!producto
                                Vector(Renglon, 4) = Str$(rstHoja!Real)
                                Vector(Renglon, 5) = "1"
                            End If
                        End If
                    
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstHoja.Close
    End If
 
  
    
    
    
    
    XParam = "'" + "AA-00000-000" + "','" _
                 + "ZZ-99999-999" + "'"
    
      
    
    spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstMovguia.RecordCount > 0 Then
    
        With rstMovguia
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    If rstMovguia!Marca <> "X" And rstMovguia!Saldo <> 0 Then
                    
                    If rstMovguia!Tipo = "T" Then
                    
                        WCantidad = rstMovguia!Cantidad
                        WMovi = rstMovguia!Movi
                        If WMovi = "S" Then
                            Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                Else
                            Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                        End If
                        
                        If WMovi = "E" Then
                            Entra = "S"
                            For dada = 1 To Renglon
                                If Vector(dada, 1) = Lote Then
                                    Entra = "N"
                                    Exit For
                                End If
                            Next dada
                            If Entra = "S" Then
                                If Lote <> 0 Then
                                    Renglon = Renglon + 1
                                    Vector(Renglon, 1) = Lote
                                    Vector(Renglon, 2) = rstMovguia!Saldo
                                    Vector(Renglon, 3) = rstMovguia!Terminado
                                    Vector(Renglon, 4) = 0
                                    Vector(Renglon, 5) = "2"
                                End If
                            End If
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
  
    Rem DADA
    Rem DADA
    Rem DADA
    Rem DADA
  
    Rem Renglon = 1
    
    Rem Vector(Renglon, 1) = "304299"
    Rem Vector(Renglon, 2) = "0.5"
    Rem Vector(Renglon, 3) = "PT-25130-100"
    Rem Vector(Renglon, 4) = "0"
    Rem Vector(Renglon, 5) = "2"
    
    

    For dada = 1 To Renglon
    
        WLote = Vector(dada, 1)
        WSaldo = Val(Vector(dada, 2))
        WTerminado = Vector(dada, 3)
        XSaldo = Val(Vector(dada, 4))
        WOrigen = Val(Vector(dada, 5))
        
        spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
            WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
            rstTerminado.Close
        End If
        
        XParam = "'" + WTerminado + "','" _
                     + WTerminado + "'"
        spEstadistica = "ListaEstadisticaDesdeHasta" + XParam
        Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEstadistica.RecordCount > 0 Then
    
            With rstEstadistica
    
                .MoveFirst
                
                If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Xmarca = IIf(IsNull(rstEstadistica!Marca), "", rstEstadistica!Marca)
                    
                    If Xmarca <> "X" Then
                    
                        aa = rstEstadistica!Numero
                        
                        WCanti1 = rstEstadistica!Canti1
                        If WCanti1 = 0 Then
                            WCanti1 = rstEstadistica!Cantidad
                        End If
                    
                        If rstEstadistica!lote1 = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - WCanti1
                                    Else
                                XSaldo = XSaldo + WCanti1
                            End If
                        End If
                        If rstEstadistica!lote2 = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - rstEstadistica!Canti2
                                    Else
                                XSaldo = XSaldo + rstEstadistica!Canti2
                            End If
                        End If
                        If rstEstadistica!lote3 = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - rstEstadistica!Canti3
                                    Else
                                XSaldo = XSaldo + rstEstadistica!Canti3
                            End If
                        End If
                        If rstEstadistica!lote4 = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - rstEstadistica!Canti4
                                    Else
                                XSaldo = XSaldo + rstEstadistica!Canti4
                            End If
                        End If
                        If rstEstadistica!lote5 = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - rstEstadistica!Canti5
                                    Else
                                XSaldo = XSaldo + rstEstadistica!Canti5
                            End If
                        End If
                        
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            ZZLote6 = Mid$(WLoteAdicional, 1, 8)
                            ZZCanti6 = Mid$(WLoteAdicional, 9, 6)
                            ZZLote7 = Mid$(WLoteAdicional, 15, 8)
                            ZZCanti7 = Mid$(WLoteAdicional, 23, 6)
                            ZZLote8 = Mid$(WLoteAdicional, 29, 8)
                            ZZCanti8 = Mid$(WLoteAdicional, 37, 6)
                            ZZLote9 = Mid$(WLoteAdicional, 43, 8)
                            ZZCanti9 = Mid$(WLoteAdicional, 51, 6)
                            ZZLote10 = Mid$(WLoteAdicional, 57, 8)
                            ZZCanti10 = Mid$(WLoteAdicional, 65, 6)
                            ZZLote11 = Mid$(WLoteAdicional, 71, 8)
                            ZZCanti11 = Mid$(WLoteAdicional, 79, 6)
                            ZZLote12 = Mid$(WLoteAdicional, 85, 8)
                            ZZCanti12 = Mid$(WLoteAdicional, 93, 6)
                                Else
                            ZZLote6 = ""
                            ZZCanti6 = "0"
                            ZZLote7 = ""
                            ZZCanti7 = "0"
                            ZZLote8 = ""
                            ZZCanti8 = "0"
                            ZZLote9 = ""
                            ZZCanti9 = "0"
                            ZZLote10 = ""
                            ZZCanti10 = "0"
                            ZZLote11 = ""
                            ZZCanti11 = "0"
                            ZZLote12 = ""
                            ZZCanti12 = "0"
                        End If
                        
                        If Val(ZZLote6) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti6)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti6)
                            End If
                        End If
                        
                        If Val(ZZLote7) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti7)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti7)
                            End If
                        End If
                        
                        If Val(ZZLote8) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti8)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti8)
                            End If
                        End If
                        
                        If Val(ZZLote9) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti9)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti9)
                            End If
                        End If
                        
                        If Val(ZZLote10) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti10)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti10)
                            End If
                        End If
                        
                        If Val(ZZLote11) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti11)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti11)
                            End If
                        End If
                        
                        If Val(ZZLote12) = Val(WLote) Then
                            If rstEstadistica!Tipo = 1 Then
                                XSaldo = XSaldo - Val(ZZCanti12)
                                    Else
                                XSaldo = XSaldo + Val(ZZCanti12)
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
            rstEstadistica.Close
        End If
    
        XParam = "'" + WTerminado + "','" _
                     + WTerminado + "'"
        spHoja = "ListaHojaTerminadoDesdeHasta" + XParam
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
                
                                If rstHoja!lote1 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti1
                                End If
                                If rstHoja!lote2 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti2
                                End If
                                If rstHoja!lote3 = Val(WLote) Then
                                    XSaldo = XSaldo - rstHoja!Canti3
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
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovvar = "ListaMovvarTerminadoDesdeHasta" + XParam
        Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovvar.RecordCount > 0 Then
    
            With rstMovvar
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovvar!Marca <> "X" Then
                
                        If rstMovvar!Tipo = "T" Then
                            WCantidad = rstMovvar!Cantidad
                            WMovi = rstMovvar!Movi
                            Lote = IIf(IsNull(rstMovvar!Lote), "0", rstMovvar!Lote)
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
                                End If
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
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovguia = "ListaMovguiaTerminadoDesdeHasta" + XParam
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
    
            With rstMovguia
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        If rstMovguia!Marca <> "X" Then
                
                        If rstMovguia!Tipo = "T" Then
                            WCantidad = rstMovguia!Cantidad
                            WMovi = rstMovguia!Movi
                            If WMovi = "S" Then
                                Lote = IIf(IsNull(rstMovguia!Partida), "0", rstMovguia!Partida)
                                    Else
                                Lote = IIf(IsNull(rstMovguia!Lote), "0", rstMovguia!Lote)
                            End If
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
                                End If
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
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spConsig = "ListaConsigTerminado" + XParam
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
                            Lote = IIf(IsNull(rstConsig!Lote), "0", rstConsig!Lote)
                            If WCantidad <> 0 Then
                                If Val(Lote) = Val(WLote) Then
                                    XSaldo = XSaldo - WCantidad
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
            rstConsig.Close
        End If
    
        XParam = "'" + WTerminado + "','" _
                 + WTerminado + "'"
        spMovlab = "ListaMovlabTerminadoDesdeHasta" + XParam
        Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovlab.RecordCount > 0 Then
    
            With rstMovlab
    
                .MoveFirst
            
                If .NoMatch = False Then
                    Do
            
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        If rstMovlab!Marca <> "X" Then
                
                        If rstMovlab!Tipo = "T" Then
                            WCantidad = rstMovlab!Cantidad
                            WMovi = rstMovlab!Movi
                            Lote = IIf(IsNull(rstMovlab!Lote), "0", rstMovlab!Lote)
                            If Val(Lote) = Val(WLote) Then
                                If WMovi = "E" Then
                                    XSaldo = XSaldo + WCantidad
                                        Else
                                    XSaldo = XSaldo - WCantidad
                                End If
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
        
        dd = WLote
        aad = WTerminado
        
        Call Redondeo(XSaldo)
        Call Redondeo(WSaldo)
        
        If XSaldo <> WSaldo Then
        
            If WOrigen = 1 Then
            
                XParam = "'" + WLote + "','" _
                             + WTerminado + "'"
                spHoja = "ListaHojaProducto " + XParam
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    WClave = rstHoja!Clave
                    ZSaldo = Str$(XSaldo)
                    WDate = Date$
                    rstHoja.Close
                                
                    XParam = "'" + WClave + "','" _
                                 + WDate + "','" _
                                 + ZSaldo + "'"
                    spHoja = "ModificaHojaSaldo " + XParam
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
                
                        Else
                        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Guia"
                ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
                ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                ZSql = ZSql + " Order by fechaord"
                spGuia = ZSql
                Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                If rstGuia.RecordCount > 0 Then
                    WClave = rstGuia!Clave
                    ZSaldo = Str$(XSaldo)
                    rstGuia.Close
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Guia SET "
                    ZSql = ZSql + " Saldo = 0"
                    ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
                    ZSql = ZSql + " and Guia.Lote = " + "'" + WLote + "'"
                    spGuia = ZSql
                    Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Guia SET "
                    ZSql = ZSql + " Saldo = " + "'" + ZSaldo + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                    spGuia = ZSql
                    Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
                End If
                            
            End If
            
        End If
    
    Next dada
    
    Next A
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()
    PrgVerilot2Auto.Hide
    Unload Me
    PrgProcVto.Show
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub
