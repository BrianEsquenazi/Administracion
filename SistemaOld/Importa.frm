VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgImporta 
   AutoRedraw      =   -1  'True
   Caption         =   "Comparacion de Importaciones/Exportaciones entre periodos"
   ClientHeight    =   4155
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   3375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   5655
      Begin VB.ComboBox Tipo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   2160
         Width           =   2175
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   8
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
         Left            =   2280
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
      Begin VB.OptionButton Impresora 
         Caption         =   "Impresora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   1215
      End
      Begin VB.OptionButton Panta 
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaII 
         Height          =   300
         Left            =   2280
         TabIndex        =   11
         Top             =   1680
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
      Begin MSMask.MaskEdBox DesdeII 
         Height          =   300
         Left            =   2280
         TabIndex        =   12
         Top             =   1320
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
      Begin VB.Label Label1 
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
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Listado"
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
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7080
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinf.rpt"
      Destination     =   1
      WindowTitle     =   "Comparacion de Importaciones/Exportaciones"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim ZCarga(100, 2) As String
Dim ZZMovi(5000, 10) As String
Dim CargaEmpresa(12, 2) As String

Private Sub Acepta_Click()

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With
    
    With rstImporta
        .Index = "Codigo"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                     Exit Do
                End If
            Loop
        End If
    End With
    
    
    
    WDesdeI = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHastaI = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    WAnoI = Right$(Desde.Text, 4)
    
    WDesdeII = Right$(DesdeII.Text, 4) + Mid$(DesdeII.Text, 4, 2) + Left$(DesdeII.Text, 2)
    WHastaII = Right$(HastaII.Text, 4) + Mid$(HastaII.Text, 4, 2) + Left$(HastaII.Text, 2)
    WAnoII = Right$(DesdeII.Text, 4)
    
    If Tipo.ListIndex = 0 Then
    
        ZLugar = 0
        Erase ZCarga
    
        spArticulo = "ListaArticulo"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstArticulo
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZZCodigo = !Codigo
                    ZZDescripcion = !Descripcion
            
                    With rstImporta
                        .Index = "Codigo"
                        .Seek "=", ZZCodigo
                        If .NoMatch Then
                            .AddNew
                            !Codigo = ZZCodigo
                            !Descripcion = ZZDescripcion
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !AnoI = WAnoI
                            !AnoII = WAnoII
                            !Titulo = "del " + Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                            !DesEmpresa = WDesEmpresa
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = ZZCodigo
                            !Descripcion = ZZDescripcion
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !AnoI = WAnoI
                            !AnoII = WAnoII
                            !Titulo = "del " + Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                            !DesEmpresa = WDesEmpresa
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstArticulo.Close
    
    
            
        XEmpresa = WEmpresa
        Erase CargaEmpresa
    
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
            CargaEmpresa(2, 1) = "0003"
            CargaEmpresa(2, 2) = "Empresa03"
            CargaEmpresa(3, 1) = "0005"
            CargaEmpresa(3, 2) = "Empresa05"
            CargaEmpresa(4, 1) = "0006"
            CargaEmpresa(4, 2) = "Empresa06"
            CargaEmpresa(5, 1) = "0007"
            CargaEmpresa(5, 2) = "Empresa07"
            CargaEmpresa(6, 1) = "0010"
            CargaEmpresa(6, 2) = "Empresa10"
            CargaEmpresa(7, 1) = "0011"
            CargaEmpresa(7, 2) = "Empresa11"
            
                Else
            
            CargaEmpresa(1, 1) = "0002"
            CargaEmpresa(1, 2) = "Empresa02"
            CargaEmpresa(2, 1) = "0004"
            CargaEmpresa(2, 2) = "Empresa04"
            CargaEmpresa(3, 1) = "0008"
            CargaEmpresa(3, 2) = "Empresa08"
            CargaEmpresa(4, 1) = "0009"
            CargaEmpresa(4, 2) = "Empresa09"
            
        End If
        
        For Cicla = 1 To 7
        
            If CargaEmpresa(Cicla, 1) <> "" Then
        
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
              
                ZZLugar = 0
                Erase ZZMovi
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Order by Informe.Articulo"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
            
                    With rstInforme
            
                        .MoveFirst
                    
                        If .NoMatch = False Then
                            Do
                    
                                If .EOF = True Then
                                    Exit Do
                                End If
                                
                                WFechaord = IIf(IsNull(!FechaOrd), "", !FechaOrd)
                                
                                
                                If WFechaord >= WDesdeI And WFechaord <= WHastaI Then
                                
                                    WInforme = rstInforme!Informe
                                    WFecha = rstInforme!Fecha
                                    WFechaord = rstInforme!FechaOrd
                                    WArticulo = rstInforme!Articulo
                                    WCantidad = rstInforme!Cantidad
                                    WOrden = rstInforme!Orden
                                
                                    ZZLugar = ZZLugar + 1
                                    
                                    ZZMovi(ZZLugar, 1) = Str$(WInforme)
                                    ZZMovi(ZZLugar, 2) = WFecha
                                    ZZMovi(ZZLugar, 3) = WFechaord
                                    ZZMovi(ZZLugar, 4) = WArticulo
                                    ZZMovi(ZZLugar, 5) = Str$(WCantidad)
                                    ZZMovi(ZZLugar, 6) = Str$(WOrden)
                                    ZZMovi(ZZLugar, 7) = "1"
                                    
                                End If
                        
                                If WFechaord >= WDesdeII And WFechaord <= WHastaII Then
                                
                                    WInforme = rstInforme!Informe
                                    WFecha = rstInforme!Fecha
                                    WFechaord = rstInforme!FechaOrd
                                    WArticulo = rstInforme!Articulo
                                    WCantidad = rstInforme!Cantidad
                                    WOrden = rstInforme!Orden
                                
                                    ZZLugar = ZZLugar + 1
                                    
                                    ZZMovi(ZZLugar, 1) = Str$(WInforme)
                                    ZZMovi(ZZLugar, 2) = WFecha
                                    ZZMovi(ZZLugar, 3) = WFechaord
                                    ZZMovi(ZZLugar, 4) = WArticulo
                                    ZZMovi(ZZLugar, 5) = Str$(WCantidad)
                                    ZZMovi(ZZLugar, 6) = Str$(WOrden)
                                    ZZMovi(ZZLugar, 7) = "2"
                                    
                                End If
                        
                                .MoveNext
                            
                                If .EOF = True Then
                                    Exit Do
                                End If
                        
                            Loop
                        End If
                    End With
                    
                    rstInforme.Close
                
                End If
                
                For ZZCiclo = 1 To ZZLugar
            
                    WInforme = Val(ZZMovi(ZZCiclo, 1))
                    WFecha = ZZMovi(ZZCiclo, 2)
                    WFechaord = ZZMovi(ZZCiclo, 3)
                    WArticulo = ZZMovi(ZZCiclo, 4)
                    WCantidad = Val(ZZMovi(ZZCiclo, 5))
                    WOrden = ZZMovi(ZZCiclo, 6)
                    WLugar = ZZMovi(ZZCiclo, 7)
                    WMoneda = 0
                    WTipo = 0
                    WCosto = 0
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Orden"
                    ZSql = ZSql + " Where Orden = " + "'" + WOrden + "'"
                    ZSql = ZSql + " and Articulo = " + "'" + WArticulo + "'"
                    spOrden = ZSql
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        WTipo = rstOrden!Tipo
                        WCosto = rstOrden!Precio
                        WMoneda = rstOrden!Moneda
                        rstOrden.Close
                    End If
                    
                    WImporte = WCantidad * WCosto
                    
                    If WTipo = 1 Then
                    
                        With rstImporta
                            .Index = "Codigo"
                            .Seek "=", WArticulo
                            If .NoMatch Then
                                .AddNew
                                
                                !Codigo = WArticulo
                                !Descripcion = ""
                                !Importe1 = 0
                                !Importe2 = 0
                                !Importe3 = 0
                                !Importe4 = 0
                                !Importe5 = 0
                                !Importe6 = 0
                                
                                If Val(WLugar) = 1 Then
                                    !Importe1 = !Importe1 + WCantidad
                                    If Val(WMoneda) = 0 Then
                                        !Importe2 = !Importe2 + WImporte
                                            Else
                                        !Importe3 = !Importe3 + WImporte
                                    End If
                                End If
                                
                                If Val(WLugar) = 2 Then
                                    !Importe4 = !Importe4 + WCantidad
                                    If Val(WMoneda) = 0 Then
                                        !Importe5 = !Importe5 + WImporte
                                            Else
                                        !Importe6 = !Importe6 + WImporte
                                    End If
                                End If
                                
                                .Update
                                .Bookmark = .LastModified
                                    Else
                                .Edit
                                
                                If Val(WLugar) = 1 Then
                                    !Importe1 = !Importe1 + WCantidad
                                    If Val(WMoneda) = 0 Then
                                        !Importe2 = !Importe2 + WImporte
                                            Else
                                        !Importe3 = !Importe3 + WImporte
                                    End If
                                End If
                                
                                If Val(WLugar) = 2 Then
                                    !Importe4 = !Importe4 + WCantidad
                                    If Val(WMoneda) = 0 Then
                                        !Importe5 = !Importe5 + WImporte
                                            Else
                                        !Importe6 = !Importe6 + WImporte
                                    End If
                                End If
                                
                                .Update
                                .Bookmark = .LastModified
                            End If
                        End With
                    
                    End If
                    
                Next ZZCiclo
                      
            End If
                
        Next Cicla
            
        With rstImporta
            .Index = "Codigo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    If !Importe1 = 0 And !Importe2 = 0 And !Importe3 = 0 And !Importe4 = 0 And !Importe5 = 0 And !Importe6 = 0 Then
                        .Delete
                    End If
                    .MoveNext
                    If .EOF = True Then
                         Exit Do
                    End If
                Loop
            End If
        End With
            
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
        
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            Listado.DataFiles(0) = "0001" + "Auxi.mdb"
                Else
            Listado.DataFiles(0) = "0002" + "Auxi.mdb"
        End If
        
        If Impresora.Value = True Then
            Listado.Destination = 1
                Else
            Listado.Destination = 0
        End If
        
        Listado.ReportFileName = "ImportaI.rpt"
    
        Listado.Action = 1
        
        
            Else
    
    
        ZLugar = 0
        Erase ZCarga
    
        spTerminado = "ListaTerminado"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
        With rstTerminado
            .MoveFirst
            Do
                If .EOF = False Then
            
                    ZZCodigo = !Codigo
                    ZZDescripcion = !Descripcion
            
                    With rstImporta
                        .Index = "Codigo"
                        .Seek "=", ZZCodigo
                        If .NoMatch Then
                            .AddNew
                            !Codigo = ZZCodigo
                            !Descripcion = ZZDescripcion
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !AnoI = WAnoI
                            !AnoII = WAnoII
                            !Titulo = "del " + Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                            !DesEmpresa = WDesEmpresa
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            !Codigo = ZZCodigo
                            !Descripcion = ZZDescripcion
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            !AnoI = WAnoI
                            !AnoII = WAnoII
                            !Titulo = "del " + Left$(Desde.Text, 5) + " al " + Left$(Hasta.Text, 5)
                            !DesEmpresa = WDesEmpresa
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTerminado.Close
    
    
            
        XEmpresa = WEmpresa
        Erase CargaEmpresa
    
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            CargaEmpresa(1, 1) = "0001"
            CargaEmpresa(1, 2) = "Empresa01"
                Else
            CargaEmpresa(1, 1) = "0008"
            CargaEmpresa(1, 2) = "Empresa08"
        End If
        
        For Cicla = 1 To 1
        
            If CargaEmpresa(Cicla, 1) <> "" Then
        
                WEmpresa = CargaEmpresa(Cicla, 1)
                txtOdbc = CargaEmpresa(Cicla, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
              
                ZZLugar = 0
                Erase ZZMovi
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Estadistica"
                ZSql = ZSql + " Where Estadistica.Numero >= " + "'" + "800000" + "'"
                ZSql = ZSql + " Order by Estadistica.Articulo"
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
                                
                                If !Cliente <> "A00066" Then
                                
                                WFechaord = IIf(IsNull(!OrdFecha), "", !OrdFecha)
                                
                                If WFechaord >= WDesdeI And WFechaord <= WHastaI Then
                                
                                    If !Numero >= 800000 Then
                                    
                                        WNumero = !Numero
                                        WFecha = !Fecha
                                        WFechaord = !OrdFecha
                                        WArticulo = !Articulo
                                        WCantidad = !Cantidad
                                        WPrecio = !Importe
                                        WPrecioUs = !ImporteUS
                                    
                                        ZZLugar = ZZLugar + 1
                                        
                                        ZZMovi(ZZLugar, 1) = Str$(WInforme)
                                        ZZMovi(ZZLugar, 2) = WFecha
                                        ZZMovi(ZZLugar, 3) = WFechaord
                                        ZZMovi(ZZLugar, 4) = WArticulo
                                        ZZMovi(ZZLugar, 5) = Str$(WCantidad)
                                        ZZMovi(ZZLugar, 6) = Str$(WPrecio)
                                        ZZMovi(ZZLugar, 7) = Str$(WPrecioUs)
                                        ZZMovi(ZZLugar, 8) = "1"
                                        
                                    End If
                                    
                                End If
                        
                                If WFechaord >= WDesdeII And WFechaord <= WHastaII Then
                                
                                    If !Numero >= 800000 Then
                                    
                                        WNumero = !Numero
                                        WFecha = !Fecha
                                        WFechaord = !OrdFecha
                                        WArticulo = !Articulo
                                        WCantidad = !Cantidad
                                        WPrecio = !Importe
                                        WPrecioUs = !ImporteUS
                                    
                                        ZZLugar = ZZLugar + 1
                                        
                                        ZZMovi(ZZLugar, 1) = Str$(WNumero)
                                        ZZMovi(ZZLugar, 2) = WFecha
                                        ZZMovi(ZZLugar, 3) = WFechaord
                                        ZZMovi(ZZLugar, 4) = WArticulo
                                        ZZMovi(ZZLugar, 5) = Str$(WCantidad)
                                        ZZMovi(ZZLugar, 6) = Str$(WPrecio)
                                        ZZMovi(ZZLugar, 7) = Str$(WPrecioUs)
                                        ZZMovi(ZZLugar, 8) = "2"
                                        
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
                
                For ZZCiclo = 1 To ZZLugar
            
                    WNumero = Val(ZZMovi(ZZCiclo, 1))
                    WFecha = ZZMovi(ZZCiclo, 2)
                    WFechaord = ZZMovi(ZZCiclo, 3)
                    WArticulo = ZZMovi(ZZCiclo, 4)
                    WCantidad = Val(ZZMovi(ZZCiclo, 5))
                    WImporte = Val(ZZMovi(ZZCiclo, 6))
                    WImporteUs = Val(ZZMovi(ZZCiclo, 7))
                    WLugar = ZZMovi(ZZCiclo, 8)
                    
                    With rstImporta
                        .Index = "Codigo"
                        .Seek "=", WArticulo
                        If .NoMatch Then
                            .AddNew
                            
                            !Codigo = WArticulo
                            !Descripcion = ""
                            !Importe1 = 0
                            !Importe2 = 0
                            !Importe3 = 0
                            !Importe4 = 0
                            !Importe5 = 0
                            !Importe6 = 0
                            
                            If Val(WLugar) = 1 Then
                                !Importe1 = !Importe1 + WCantidad
                                !Importe2 = !Importe2 + WImporteUs
                                !Importe3 = !Importe3 + WImporte
                            End If
                            
                            If Val(WLugar) = 2 Then
                                !Importe4 = !Importe4 + WCantidad
                                !Importe5 = !Importe5 + WImporteUs
                                !Importe6 = !Importe6 + WImporte
                            End If
                            
                            .Update
                            .Bookmark = .LastModified
                                Else
                            .Edit
                            
                            If Val(WLugar) = 1 Then
                                !Importe1 = !Importe1 + WCantidad
                                !Importe2 = !Importe2 + WImporteUs
                                !Importe3 = !Importe3 + WImporte
                            End If
                            
                            If Val(WLugar) = 2 Then
                                !Importe4 = !Importe4 + WCantidad
                                !Importe5 = !Importe5 + WImporteUs
                                !Importe6 = !Importe6 + WImporte
                            End If
                            
                            .Update
                            .Bookmark = .LastModified
                        End If
                    End With
                    
                Next ZZCiclo
                      
            End If
                
        Next Cicla
            
        With rstImporta
            .Index = "Codigo"
            .Seek ">=", ""
            If .NoMatch = False Then
                Do
                    If !Importe1 = 0 And !Importe2 = 0 And !Importe3 = 0 And !Importe4 = 0 And !Importe5 = 0 And !Importe6 = 0 Then
                        .Delete
                    End If
                    .MoveNext
                    If .EOF = True Then
                         Exit Do
                    End If
                Loop
            End If
        End With
            
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
        
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
            Listado.DataFiles(0) = "0001" + "Auxi.mdb"
                Else
            Listado.DataFiles(0) = "0002" + "Auxi.mdb"
        End If
        
        If Impresora.Value = True Then
            Listado.Destination = 1
                Else
            Listado.Destination = 0
        End If
        
        Listado.ReportFileName = "ImportaII.rpt"
    
        Listado.Action = 1
            
    End If
    
    
    
End Sub

Private Sub Cancela_click()
    PrgImporta.Hide
    Unload Me
    Menu.Show
End Sub


Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Importa
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub


Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeII.SetFocus
    End If
End Sub

Private Sub DesdeII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaII.SetFocus
    End If
End Sub

Private Sub HastaUII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub


Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Importaciones"
    Tipo.AddItem "Exportaciones"
    
    Tipo.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgImporta.Caption = "Comparacion de Importaciones/Exportaciones entre periodos :  " + !Nombre
        End If
    End With
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    DesdeII.Text = "  /  /    "
    HastaII.Text = "  /  /    "
    Panta.Value = False
    
    Impresora.Value = True
    Frame2.Visible = True
End Sub

