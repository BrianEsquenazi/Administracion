VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaImpoFecha 
   Caption         =   "Listado de Partidas de Importaciones Valorizadas a Fecha"
   ClientHeight    =   2520
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   2520
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1800
         TabIndex        =   0
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label4 
         Caption         =   "Desde Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Hasta Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7680
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "ListaImpoFecha.rpt"
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
End
Attribute VB_Name = "PrgListaImpoFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstLoteFecha As Recordset
Dim spLoteFecha As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim Empe(12, 10) As String
Dim Vector(5000, 10) As String
Dim XLote(12, 2) As String

Private Sub Acepta_Click()

    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)

    ZSql = "DELETE LoteFecha"
    spLoteFecha = ZSql
    Set rstLoteFecha = db.OpenRecordset(spLoteFecha, dbOpenSnapshot, dbSQLPassThrough)

    XEmpresa = WEmpresa
    
    If Val(XEmpresa) = 1 Or Val(XEmpresa) = 3 Or Val(XEmpresa) = 5 Or Val(XEmpresa) = 6 Or Val(XEmpresa) = 7 Or Val(XEmpresa) = 10 Or Val(XEmpresa) = 11 Then
        Empe(1, 1) = "0001"
        Empe(1, 2) = "Empresa01"
        Empe(2, 1) = "0003"
        Empe(2, 2) = "Empresa03"
        Empe(3, 1) = "0005"
        Empe(3, 2) = "Empresa05"
        Empe(4, 1) = "0006"
        Empe(4, 2) = "Empresa06"
        Empe(5, 1) = "0007"
        Empe(5, 2) = "Empresa07"
        Empe(6, 1) = "0010"
        Empe(6, 2) = "Empresa10"
        Empe(7, 1) = "0011"
        Empe(7, 2) = "Empresa11"
        XHasta = 7
            Else
        Empe(1, 1) = "0002"
        Empe(1, 2) = "Empresa02"
        Empe(2, 1) = "0004"
        Empe(2, 2) = "Empresa04"
        Empe(3, 1) = "0008"
        Empe(3, 2) = "Empresa08"
        Empe(4, 1) = "0009"
        Empe(4, 2) = "Empresa09"
        XHasta = 4
    End If
    
    
    Erase Vector
    Lugar = 0
    
    
    For a = 1 To XHasta
    
        WEmpresa = Empe(a, 1)
        txtOdbc = Empe(a, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        WRenglon = "01"
        
        ZSql = ""
        ZSql = ZSql + "Select *, Orden.Tipo as [WTipoOrden], Orden.Carpeta as [WCarpeta]"
        ZSql = ZSql + " FROM Laudo, Orden"
        ZSql = ZSql + " Where Laudo.Orden = Orden.Orden"
        ZSql = ZSql + " and Orden.Renglon = " + "'" + WRenglon + "'"
        ZSql = ZSql + " and Laudo.FechaOrd >= " + "'" + WDesde + "'"
        ZSql = ZSql + " and Laudo.FechaOrd <= " + "'" + WHasta + "'"
        ZSql = ZSql + " Order by Laudo.Clave"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
        
            With rstLaudo
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        If rstLaudo!WTipoOrden = 1 Then
                        If Left$(rstLaudo!Articulo, 2) <> "DY" And Left$(rstLaudo!Articulo, 2) <> "DS" And Left$(rstLaudo!Articulo, 2) <> "DQ" Then
                        
                            Entra = "S"
                            
                            For ZCicla = 1 To Lugar
                                If Left$(rstLaudo!Articulo, 2) = "DY" Or Left$(rstLaudo!Articulo, 2) = "DS" Or Left$(rstLaudo!Articulo, 2) = "DQ" Then
                                    If Vector(ZCicla, 7) = rstLaudo!PartiOri And Vector(ZCicla, 4) = rstLaudo!Articulo Then
                                        If rstLaudo!Liberadaant > 0 Then
                                            Vector(ZCicla, 5) = Str$(Val(Vector(ZCicla, 5)) + rstLaudo!Liberadaant)
                                                Else
                                            Vector(ZCicla, 5) = Str$(Val(Vector(ZCicla, 5)) + rstLaudo!Liberada)
                                        End If
                                        Entra = "N"
                                        Exit For
                                    End If
                                End If
                            Next ZCicla
                            
                            If Entra = "S" Then
                                Lugar = Lugar + 1
                                Vector(Lugar, 1) = rstLaudo!Proveedor
                                Vector(Lugar, 2) = rstLaudo!WCarpeta
                                Vector(Lugar, 3) = rstLaudo!Fecha
                                Vector(Lugar, 4) = rstLaudo!Articulo
                                If rstLaudo!Liberadaant > 0 Then
                                    Vector(Lugar, 5) = Str$(rstLaudo!Liberadaant)
                                        Else
                                    Vector(Lugar, 5) = Str$(rstLaudo!Liberada)
                                End If
                                Vector(Lugar, 6) = rstLaudo!Lote
                                Rem Vector(Lugar, 7) = rstLaudo!PartiOri
                                Vector(Lugar, 9) = WEmpresa
                            End If
                            
                        End If
                        End If
                
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            
            rstLaudo.Close
        
        End If
    
    Next a
    
    
    For Ciclo = 1 To Lugar
    
    
    
    
        WWArticulo = Vector(Ciclo, 4)
        WWLote = Vector(Ciclo, 6)
        WWSalida = 0
    
        For a = 1 To XHasta
    
    
    
            WEmpresa = Empe(a, 1)
            txtOdbc = Empe(a, 2)
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            
            
            
            XParam = "'" + WWArticulo + "','" _
                         + WWArticulo + "'"
            spHoja = "ListaHojaArticuloDesdeHasta" + XParam
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            If rstHoja.RecordCount > 0 Then
    
                With rstHoja
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
            
                            If .EOF = True Then
                                Exit Do
                            End If
                            
                            If rstHoja!FechaOrd >= WDesde And rstHoja!FechaOrd <= WHasta Then
                            
                                XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                                XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                                XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                                XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                                XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                                XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                            
                                If Val(XLote(1, 1)) = 0 Then
                                    XLote(1, 1) = rstHoja!Lote
                                    XLote(1, 2) = rstHoja!Cantidad
                                End If
                        
                                For Da = 1 To 3
                        
                                    If XLote(Da, 2) = "" Then
                                        XLote(Da, 2) = "0"
                                    End If
                        
                                    ZCanti = XLote(Da, 2)
                                    ZLote = XLote(Da, 1)
                                    If Val(WWLote) = Val(ZLote) And ZCanti <> 0 Then
                                        WWSalida = WWSalida + ZCanti
                                    End If
                                
                                Next Da
                            
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
    
    
    
    
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Movvar"
            ZSql = ZSql + " Where Movvar.Articulo = " + "'" + WWArticulo + "'"
            ZSql = ZSql + " and Movvar.Lote = " + "'" + WWLote + "'"
            spMovvar = ZSql
            Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovvar.RecordCount > 0 Then

                With rstMovvar
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
            
                            If .EOF = True Then
                                Exit Do
                            End If
                
                            If rstMovvar!FechaOrd >= WDesde And rstMovvar!FechaOrd <= WHasta Then
                                If rstMovvar!Movi = "E" Then
                                    WWSalida = WWSalida - rstMovvar!Cantidad
                                        Else
                                    WWSalida = WWSalida + rstMovvar!Cantidad
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
    
    
    
    
    
    
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Movlab"
            ZSql = ZSql + " Where Movlab.Articulo = " + "'" + WWArticulo + "'"
            ZSql = ZSql + " and Movlab.Lote = " + "'" + WWLote + "'"
            spMovlab = ZSql
            Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
            If rstMovlab.RecordCount > 0 Then

                With rstMovlab
    
                    .MoveFirst
            
                    If .NoMatch = False Then
                        Do
            
                            If .EOF = True Then
                                Exit Do
                            End If
                
                            If rstMovlab!FechaOrd >= WDesde And rstMovlab!FechaOrd <= WHasta Then
                                If rstMovlab!Movi = "E" Then
                                    WWSalida = WWSalida - rstMovlab!Cantidad
                                        Else
                                    WWSalida = WWSalida + rstMovlab!Cantidad
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
            
            
            
            
            
            If Left$(WWArticulo, 2) = "DY" Or Left$(WWArticulo, 2) = "DS" Or Left$(WWArticulo, 2) = "DQ" Then
            
                WArticuloDy = Left$(WWArticulo, 3) + "00" + Right$(WWArticulo, 7)
    
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WArticuloDy + "'"
                spEstadistica = ZSql
                Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEstadistica.RecordCount > 0 Then

                    With rstEstadistica
    
                        .MoveFirst
            
                        If .NoMatch = False Then
                            Do
            
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                                If rstEstadistica!OrdFecha >= WDesde And rstEstadistica!OrdFecha <= WHasta Then
                                
                                    Erase XLote
                            
                                    ZLote1 = IIf(IsNull(rstEstadistica!lote1), "0", rstEstadistica!lote1)
                                    ZCanti1 = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                                    ZLote2 = IIf(IsNull(rstEstadistica!lote2), "0", rstEstadistica!lote2)
                                    ZCanti2 = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                                    ZLote3 = IIf(IsNull(rstEstadistica!lote3), "0", rstEstadistica!lote3)
                                    ZCanti3 = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                                    ZLote4 = IIf(IsNull(rstEstadistica!lote4), "0", rstEstadistica!lote4)
                                    ZCanti4 = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                                    ZLote5 = IIf(IsNull(rstEstadistica!lote5), "0", rstEstadistica!lote5)
                                    ZCanti5 = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                                    
                                    XLote(1, 1) = Str$(ZLote1)
                                    XLote(1, 2) = Str$(ZCanti1)
                                    XLote(2, 1) = Str$(ZLote2)
                                    XLote(2, 2) = Str$(ZCanti2)
                                    XLote(3, 1) = Str$(ZLote3)
                                    XLote(3, 2) = Str$(ZCanti3)
                                    XLote(4, 1) = Str$(ZLote4)
                                    XLote(4, 2) = Str$(ZCanti4)
                                    XLote(5, 1) = Str$(ZLote5)
                                    XLote(5, 2) = Str$(ZCanti5)
                                
                                    WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                                    
                                    If Len(Trim(WLoteAdicional)) = 98 Then
                                        XLote(6, 1) = Mid$(WLoteAdicional, 1, 8)
                                        XLote(6, 2) = Mid$(WLoteAdicional, 9, 6)
                                        XLote(7, 1) = Mid$(WLoteAdicional, 15, 8)
                                        XLote(7, 2) = Mid$(WLoteAdicional, 23, 6)
                                        XLote(8, 1) = Mid$(WLoteAdicional, 29, 8)
                                        XLote(8, 2) = Mid$(WLoteAdicional, 37, 6)
                                        XLote(9, 1) = Mid$(WLoteAdicional, 43, 8)
                                        XLote(9, 2) = Mid$(WLoteAdicional, 51, 6)
                                        XLote(10, 1) = Mid$(WLoteAdicional, 57, 8)
                                        XLote(10, 2) = Mid$(WLoteAdicional, 65, 6)
                                        XLote(11, 1) = Mid$(WLoteAdicional, 71, 8)
                                        XLote(11, 2) = Mid$(WLoteAdicional, 79, 6)
                                        XLote(12, 1) = Mid$(WLoteAdicional, 85, 8)
                                        XLote(12, 2) = Mid$(WLoteAdicional, 93, 6)
                                            Else
                                        XLote(6, 1) = "0"
                                        XLote(6, 2) = "0"
                                        XLote(7, 1) = "0"
                                        XLote(7, 2) = "0"
                                        XLote(8, 1) = "0"
                                        XLote(8, 2) = "0"
                                        XLote(9, 1) = "0"
                                        XLote(9, 2) = "0"
                                        XLote(10, 1) = "0"
                                        XLote(10, 2) = "0"
                                        XLote(11, 1) = "0"
                                        XLote(11, 2) = "0"
                                        XLote(12, 1) = "0"
                                        XLote(12, 2) = "0"
                                    End If
                                
                                    For ZCiclo = 1 To 12
                                    
                                        ZZLote = Val(XLote(ZCiclo, 1))
                                        ZZCanti = Val(XLote(ZCiclo, 2))
                                                            
                                        If ZZCanti <> 0 And ZZLote = Val(WWLote) Then
                                            WCantidad = ZZCanti
                                            If rstEstadistica!Tipo = 1 Then
                                                WWSalida = WWSalida + ZZCanti
                                                    Else
                                                WWSalida = WWSalida - ZZCanti
                                            End If
                                        End If
                                
                                    Next ZCiclo
                                
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
                
            End If
            
        Next a
        
        ZZProve = Vector(Ciclo, 1)
        ZZCarpeta = Vector(Ciclo, 2)
        ZZFecha = Vector(Ciclo, 3)
        ZZArticulo = Vector(Ciclo, 4)
        ZZLote = Vector(Ciclo, 6)
        ZZPartiOri = Vector(Ciclo, 7)
        ZZEmpresa = Vector(Ciclo, 9)
        
        ZZCantidad = Vector(Ciclo, 5)
        Vector(Ciclo, 8) = Str$(WWSalida)
        
    Next Ciclo
    
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
            Case 10
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case 11
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
    End Select


    For Ciclo = 1 To Lugar

        ZZProve = Vector(Ciclo, 1)
        ZZCarpeta = Vector(Ciclo, 2)
        ZZFecha = Vector(Ciclo, 3)
        ZZArticulo = Vector(Ciclo, 4)
        ZZCantidad = Vector(Ciclo, 5)
        ZZLote = Vector(Ciclo, 6)
        ZZPartiOri = Vector(Ciclo, 7)
        ZZSalida = Vector(Ciclo, 8)
        ZZEmpresa = Vector(Ciclo, 9)
        ZZCosto = ""
        
        spArticulo = "ConsultaArticulo " + "'" + ZZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZZCosto = Str$(rstArticulo!Costo1)
            rstArticulo.Close
        End If
        
        If Val(ZZCantidad) < Val(ZZSalida) Then
            ZZSalida = ZZCantidad
        End If
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO LoteFecha ("
        ZSql = ZSql + "Proveedor ,"
        ZSql = ZSql + "Carpeta ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Articulo ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "PartiOri ,"
        ZSql = ZSql + "Empresa ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Salida ,"
        ZSql = ZSql + "Costo )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZZProve + "',"
        ZSql = ZSql + "'" + ZZCarpeta + "',"
        ZSql = ZSql + "'" + ZZFecha + "',"
        ZSql = ZSql + "'" + ZZArticulo + "',"
        ZSql = ZSql + "'" + ZZLote + "',"
        ZSql = ZSql + "'" + ZZPartiOri + "',"
        ZSql = ZSql + "'" + ZZEmpresa + "',"
        ZSql = ZSql + "'" + ZZCantidad + "',"
        ZSql = ZSql + "'" + ZZSalida + "',"
        ZSql = ZSql + "'" + ZZCosto + "')"
        
        spLoteFecha = ZSql
        Set rstLoteFecha = db.OpenRecordset(spLoteFecha, dbOpenSnapshot, dbSQLPassThrough)
            
    Next Ciclo
    
    

    Listado.WindowTitle = "Listado del Estado de las Importaciones"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    WDesdeProveedor = "0"
    WHastaProveedor = "999999999999"
    
    Listado.GroupSelectionFormula = "{LoteFecha.Proveedor} in " + Chr$(34) + WDesdeProveedor + Chr$(34) + " to " + Chr$(34) + WHastaProveedor + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT LoteFecha.Proveedor, LoteFecha.Carpeta, LoteFecha.Fecha, LoteFecha.Articulo, LoteFecha.Lote, LoteFecha.Cantidad, LoteFecha.Salida, LoteFecha.Costo," _
                + "Proveedor.Nombre " _
                + "From " _
                + DSQ + ".dbo.LoteFecha LoteFecha, " _
                + DSQ + ".dbo.Proveedor Proveedor " _
                + "Where " _
                + "LoteFecha.Proveedor = Proveedor.Proveedor AND " _
                + "LoteFecha.Proveedor >= '0' AND " _
                + "LoteFecha.Proveedor <= '99999999999'"
    
    Rem Listado.DataFiles(1) = WEmpresa + "auxi.mdb"
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    With rstEmpresa
        .Close
    End With
    PrgListaImpoFecha.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

