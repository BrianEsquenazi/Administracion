VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgAnalisisPagoPrv 
   AutoRedraw      =   -1  'True
   Caption         =   "Analisis de Pago de Facturas de Proveedores"
   ClientHeight    =   7365
   ClientLeft      =   450
   ClientTop       =   825
   ClientWidth     =   11100
   LinkTopic       =   "Form2"
   ScaleHeight     =   7365
   ScaleWidth      =   11100
   Begin VB.TextBox WFactura 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   11
      Top             =   6840
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid IngresoDatos 
      Height          =   4335
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7646
      _Version        =   327680
      Rows            =   1000
      Cols            =   3
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   6375
      Begin VB.CommandButton Limpia 
         Caption         =   "Limpia "
         Height          =   345
         Left            =   4680
         TabIndex        =   15
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         Height          =   345
         Left            =   4680
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox Proveedor 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         Height          =   375
         Left            =   960
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   4680
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label DesProveedor 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Sedronar.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Cuenta Corriente de Proveedores"
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
      Left            =   7200
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      Height          =   6495
      ItemData        =   "AnalisisPagoPrv.frx":0000
      Left            =   6840
      List            =   "AnalisisPagoPrv.frx":0007
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Ingreso de Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   6840
      Width           =   1935
   End
End
Attribute VB_Name = "PrgAnalisisPagoPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim XParam As String
Dim Vector(100, 2) As String
Dim ZZOrden(100, 5) As String
Dim rstPagos As Recordset
Dim spPagos As String
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstCtaCtePrv As Recordset
Dim spCtaCtePrv As String
Dim rstAnalisisFactura As Recordset
Dim spAnalisisFactura As String

Private Sub Acepta_Click()

    Listado.WindowTitle = "Analisis de Pago de Facturas de Proveedores"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    ZSql = ""
    ZSql = ZSql + "DELETE AnalisisFactura"
    spAnalisisFactura = ZSql
    Set rstAnalisisFactura = db.OpenRecordset(spAnalisisFactura, dbOpenSnapshot, dbSQLPassThrough)
    
    For A = 1 To 999
    
        WFactura = IngresoDatos.TextMatrix(A, 1)
        WImporteFactura = IngresoDatos.TextMatrix(A, 2)
        WFechaFactura = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCtePrv"
        ZSql = ZSql + " Where CtaCtePrv.Proveedor = " + "'" + Proveedor.Text + "'"
        ZSql = ZSql + " and CtaCtePrv.Numero = " + "'" + WFactura + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCtePrv.RecordCount > 0 Then
            WFechaFactura = rstCtaCtePrv!Fecha
            rstCtaCtePrv.Close
        End If
        
        
        If Val(WFactura) <> 0 Then
        
            Erase Vector
            ZLugar = 0
            
            Auxi1 = WFactura
            Call Ceros(Auxi1, 8)
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Pagos"
            ZSql = ZSql + " Where Pagos.Proveedor = " + "'" + Proveedor.Text + "'"
            ZSql = ZSql + " and Pagos.Numero1 = " + "'" + Auxi1 + "'"
            spPagos = ZSql
            Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
            If rstPagos.RecordCount > 0 Then
            
                With rstPagos
    
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                        
                            ZLugar = ZLugar + 1
                            Vector(ZLugar, 1) = rstPagos!Orden
                            Vector(ZLugar, 2) = rstPagos!Fecha
                    
                            .MoveNext
                
                            If .EOF = True Then
                                Exit Do
                            End If
                
                        Loop
                    End If
        
                End With
                rstPagos.Close
                
            End If
            
            If ZLugar > 0 Then
            
                ZZLugar = 0
                Erase ZZOrden
            
                For ZCiclo = 1 To ZLugar
                
                    ZOrden = Vector(ZCiclo, 1)
                    ZFecha = Vector(ZCiclo, 2)
                    Auxi1 = ZOrden
                    Call Ceros(Auxi1, 6)
                    
                    Erase ZZOrden
                    ZZLugar = 0
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Pagos"
                    ZSql = ZSql + " Where Pagos.Orden = " + "'" + ZOrden + "'"
                    ZSql = ZSql + " and Pagos.Importe2 <> 0"
                    spPagos = ZSql
                    Set rstPagos = db.OpenRecordset(spPagos, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPagos.RecordCount > 0 Then
            
                        With rstPagos
    
                            .MoveFirst
                            If .NoMatch = False Then
                                Do
                        
                                    ZZLugar = ZZLugar + 1
                                    ZZOrden(ZZLugar, 1) = rstPagos!Tipo2
                                    ZZOrden(ZZLugar, 2) = rstPagos!Numero2
                                    ZZOrden(ZZLugar, 3) = rstPagos!Observaciones2
                                    ZZOrden(ZZLugar, 4) = rstPagos!Fecha2
                                    ZZOrden(ZZLugar, 5) = Str$(rstPagos!Importe2)
                    
                                    .MoveNext
                
                                    If .EOF = True Then
                                        Exit Do
                                    End If
                        
                                Loop
                            End If
                
                        End With
                        rstPagos.Close
                
                    End If
                    
                    
                    For WCiclo = 1 To ZZLugar
                    
                        XXProveedor = Proveedor.Text
                        XXDesProveedor = DesProveedor.Caption
                        XXFactura = WFactura
                        XXFechaFactura = WFechaFactura
                        XXImporte = WImporteFactura
                        XXOrden = ZOrden
                        XXFecha = ZFecha
                        XXTipo = ZZOrden(WCiclo, 1)
                        XXNumero = ZZOrden(WCiclo, 2)
                        XXBanco = ZZOrden(WCiclo, 3)
                        XXFechaII = ZZOrden(WCiclo, 4)
                        XXImporteII = ZZOrden(WCiclo, 5)
                        
                        ZSql = ""
                        ZSql = ZSql + "INSERT INTO AnalisisFactura ("
                        ZSql = ZSql + "Proveedor ,"
                        ZSql = ZSql + "DesProveedor ,"
                        ZSql = ZSql + "Factura ,"
                        ZSql = ZSql + "FechaFactura ,"
                        ZSql = ZSql + "Importe ,"
                        ZSql = ZSql + "Orden ,"
                        ZSql = ZSql + "Fecha ,"
                        ZSql = ZSql + "Tipo ,"
                        ZSql = ZSql + "Numero ,"
                        ZSql = ZSql + "Banco ,"
                        ZSql = ZSql + "FechaII ,"
                        ZSql = ZSql + "ImporteII )"
                        ZSql = ZSql + "Values ("
                        ZSql = ZSql + "'" + XXProveedor + "',"
                        ZSql = ZSql + "'" + XXDesProveedor + "',"
                        ZSql = ZSql + "'" + XXFactura + "',"
                        ZSql = ZSql + "'" + XXFechaFactura + "',"
                        ZSql = ZSql + "'" + XXImporte + "',"
                        ZSql = ZSql + "'" + XXOrden + "',"
                        ZSql = ZSql + "'" + XXFecha + "',"
                        ZSql = ZSql + "'" + XXTipo + "',"
                        ZSql = ZSql + "'" + XXNumero + "',"
                        ZSql = ZSql + "'" + XXBanco + "',"
                        ZSql = ZSql + "'" + XXFechaII + "',"
                        ZSql = ZSql + "'" + XXImporteII + "')"
            
                        spAnalisisFactura = ZSql
                        Set rstAnalisisFactura = db.OpenRecordset(spAnalisisFactura, dbOpenSnapshot, dbSQLPassThrough)
                    
                    Next WCiclo
                    
                Next ZCiclo
            
                    Else
                    
                XXProveedor = Proveedor.Text
                XXDesProveedor = DesProveedor.Caption
                XXFactura = WFactura
                XXFechaFactura = WFechaFactura
                XXImporte = WImporteFactura
                XXOrden = ""
                XXFecha = ""
                XXTipo = ""
                XXNumero = ""
                XXBanco = ""
                XXFechaII = ""
                XXImporteII = ""
                        
                ZSql = ""
                ZSql = ZSql + "INSERT INTO AnalisisFactura ("
                ZSql = ZSql + "Proveedor ,"
                ZSql = ZSql + "DesProveedor ,"
                ZSql = ZSql + "Factura ,"
                ZSql = ZSql + "FechaFactura ,"
                ZSql = ZSql + "Importe ,"
                ZSql = ZSql + "Orden ,"
                ZSql = ZSql + "Fecha ,"
                ZSql = ZSql + "Tipo ,"
                ZSql = ZSql + "Numero ,"
                ZSql = ZSql + "Banco ,"
                ZSql = ZSql + "FechaII ,"
                ZSql = ZSql + "ImporteII )"
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + XXProveedor + "',"
                ZSql = ZSql + "'" + XXDesProveedor + "',"
                ZSql = ZSql + "'" + XXFactura + "',"
                ZSql = ZSql + "'" + XXFechaFactura + "',"
                ZSql = ZSql + "'" + XXImporte + "',"
                ZSql = ZSql + "'" + XXOrden + "',"
                ZSql = ZSql + "'" + XXFecha + "',"
                ZSql = ZSql + "'" + XXTipo + "',"
                ZSql = ZSql + "'" + XXNumero + "',"
                ZSql = ZSql + "'" + XXBanco + "',"
                ZSql = ZSql + "'" + XXFechaII + "',"
                ZSql = ZSql + "'" + XXImporteII + "')"
                
                spAnalisisFactura = ZSql
                Set rstAnalisisFactura = db.OpenRecordset(spAnalisisFactura, dbOpenSnapshot, dbSQLPassThrough)
                    
            End If
            
            
        End If
        
    Next A

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "AnalisisFactura.rpt"
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT AnalisisFactura.Proveedor, AnalisisFactura.Factura, AnalisisFactura.Importe, AnalisisFactura.Orden, AnalisisFactura.Fecha, AnalisisFactura.Tipo, AnalisisFactura.Numero, AnalisisFactura.Banco, AnalisisFactura.FechaII, AnalisisFactura.ImporteII, AnalisisFactura.DesProveedor, AnalisisFactura.FechaFactura " _
            + "From " _
            + DSQ + ".dbo.AnalisisFactura AnalisisFactura " _
            + "Where " _
            + "AnalisisFactura.Proveedor >= '0' AND " _
            + "AnalisisFactura.Proveedor <= '99999999999'"
    
    Listado.Connect = Connect()
    
    Listado.Action = 1
    
End Sub

Private Sub Cancela_click()
    PrgAnalisisPagoPrv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub IngresoDatos_DblClick()
    IngresoDatos.Col = 1
    IngresoDatos.Text = ""
    IngresoDatos.Col = 2
    IngresoDatos.Text = ""
    WFactura.SetFocus
End Sub



Private Sub WFactura_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Auxi1 = WFactura.Text
        Call Ceros(Auxi1, 8)
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CtaCtePrv"
        ZSql = ZSql + " Where CtaCtePrv.Proveedor = " + "'" + Proveedor.Text + "'"
        ZSql = ZSql + " and CtaCtePrv.Numero = " + "'" + Auxi1 + "'"
        spCtaCtePrv = ZSql
        Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
        If rstCtaCtePrv.RecordCount > 0 Then
            WImporte = rstCtaCtePrv!Total
            rstCtaCtePrv.Close
            For A = 1 To 1000
                If IngresoDatos.TextMatrix(A, 1) = "" Then
                    IngresoDatos.Row = A
                    IngresoDatos.Col = 1
                    IngresoDatos.Text = WFactura.Text
                    IngresoDatos.Col = 2
                    IngresoDatos.Text = Str$(WImporte)
                    IngresoDatos.Text = Pusing("###,###.##", IngresoDatos.Text)
                    WFactura.Text = ""
                    WFactura.SetFocus
                    Exit For
                End If
            Next A
        End If
        
    End If
    
End Sub

Private Sub Form_Load()

    IngresoDatos.Clear
    
    IngresoDatos.ColWidth(0) = 150
    IngresoDatos.ColWidth(1) = 2000
    IngresoDatos.ColWidth(2) = 2000
    
    IngresoDatos.Row = 0
    
    IngresoDatos.Col = 1
    IngresoDatos.Text = "Factura"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Importe"
    
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = True

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    
End Sub

Private Sub Limpia_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False

    IngresoDatos.Clear
    
    IngresoDatos.ColWidth(0) = 150
    IngresoDatos.ColWidth(1) = 2000
    IngresoDatos.ColWidth(2) = 2000
    
    IngresoDatos.Row = 0
    
    IngresoDatos.Col = 1
    IngresoDatos.Text = "Factura"
    
    IngresoDatos.Col = 2
    IngresoDatos.Text = "Importe"
    
    Panta.Value = True
    Impresora.Value = False
    Frame2.Visible = True

    IngresoDatos.Col = 1
    IngresoDatos.Row = 1
    
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    
    Proveedor.SetFocus
    
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Pantalla_Click()
    Indice = Pantalla.ListIndex
    If Len(WIndice.List(Indice)) = 11 Then
        Proveedor.Text = WIndice.List(Indice)
        Call Proveedor_KeyPress(13)
            Else
        WFactura.Text = WIndice.List(Indice)
        Call WFactura_Keypress(13)
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If Ayuda.Text <> "" Then
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Else
        spProveedor = "ListaProveedoresOrdConsulta"
    End If
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                    
                    
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
                        WIndice.AddItem IngresaItem
                        Exit For
                    End If
                Next aa
                .MoveNext
                    
                        Else
                        
                Exit Do
                
            End If
        Loop
    End With
    
    RstProveedor.Close
    
    End If
    
    End If

End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spProveedor = "ConsultaProveedores " + "'" + Proveedor.Text + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
            WFactura.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub WFactura_DblClick()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CtaCtePrv"
    ZSql = ZSql + " Where CtaCtePrv.Proveedor = " + "'" + Proveedor.Text + "'"
    ZSql = ZSql + " and CtaCtePrv.Tipo = " + "'" + "01" + "'"
    ZSql = ZSql + " Order by CtaCtePrv.ordfecha"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    If rstCtaCtePrv.RecordCount > 0 Then
        With rstCtaCtePrv
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                    IngresaItem = !Numero + "  " + !Fecha + "  " + Str$(!Total)
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Numero
                    WIndice.AddItem IngresaItem
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        rstCtaCtePrv.Close
    End If

End Sub


