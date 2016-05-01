VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHistoriaTerminado 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Verificacion de Ultimos Movimientos de P.T."
   ClientHeight    =   6180
   ClientLeft      =   2085
   ClientTop       =   1500
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   840
      TabIndex        =   6
      Top             =   120
      Width           =   5535
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2160
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   1920
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   1920
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
         Left            =   4080
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
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
         Left            =   4080
         TabIndex        =   9
         Top             =   1800
         Width           =   1215
      End
      Begin MSMask.MaskEdBox Fecha 
         Height          =   300
         Left            =   2160
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
      Begin VB.Label Label3 
         Caption         =   "Fecha"
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
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Articulo"
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
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Articulo"
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
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   6600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WHistoriaTerminado.rpt"
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
      Left            =   7080
      TabIndex        =   5
      Top             =   120
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
      Height          =   3060
      ItemData        =   "historiaterminado.frx":0000
      Left            =   120
      List            =   "historiaterminado.frx":0007
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   6840
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6840
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgHistoriaTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WTerminado As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstMovvar As Recordset
Dim spMovguia As String
Dim rstMovguia As Recordset
Dim spMovvar As String
Dim rstConsig As Recordset
Dim spConsig As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstEntdev As Recordset
Dim spEntdev As String
Dim rstHistoriaTerminado As Recordset
Dim spHistoriaTerminado As String
Dim XParam As String
Dim Vector(10000, 10) As String
Dim ZTerminado(10000) As String
Private XLote(100, 7) As String
Private WCantidad As Double
Private WSaldo As Double
Dim ZLote1 As Double
Dim ZCanti1 As Double
Dim ZLote2 As Double
Dim ZCanti2 As Double
Dim ZLote3 As Double
Dim ZCanti3 As Double
Dim ZLote4 As Double
Dim ZCanti4 As Double
Dim ZLote5 As Double
Dim ZCanti5 As Double

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Erase ZTerminado
    LugarTerminado = 0
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WTituloI = !Nombre
        End If
    End With
    
    ZSql = "DELETE HistoriaTerminado"
    spHistoriaTerminado = ZSql
    Set rstHistoriaTerminado = db.OpenRecordset(spHistoriaTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Terminado"
    ZSql = ZSql + " Where Terminado.Codigo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Terminado.Codigo <= " + "'" + Hasta.Text + "'"
    spTerminado = ZSql
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        With rstTerminado
            .MoveFirst
            If .NoMatch = False Then
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    WTerminado = rstTerminado!Codigo
                    ZStock = rstTerminado!Entradas - rstTerminado!Salidas
                    If ZStock > 0 Then
                        LugarTerminado = LugarTerminado + 1
                        ZTerminado(LugarTerminado) = WTerminado
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        rstTerminado.Close
    End If
    
    Erase Vector
    LugarVector = 0
    
    
    For Ciclo = 1 To LugarTerminado
        
        WTerminado = ZTerminado(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Hoja"
        ZSql = ZSql + " Where Hoja.Producto = " + "'" + WTerminado + "'"
        ZSql = ZSql + " and Hoja.Saldo <> 0"
        ZSql = ZSql + " and Hoja.Renglon = 1"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
    
            With rstHoja
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        LugarVector = LugarVector + 1
                        Vector(LugarVector, 1) = rstHoja!Producto
                        Vector(LugarVector, 2) = Str$(rstHoja!Real)
                        Vector(LugarVector, 3) = rstHoja!Fecha
                        Vector(LugarVector, 4) = Str$(rstHoja!Hoja)
                        Vector(LugarVector, 5) = Str$(rstHoja!Saldo)
                        Vector(LugarVector, 6) = ""
                        Vector(LugarVector, 7) = ""
                        Vector(LugarVector, 8) = rstHoja!Fecha
                        Vector(LugarVector, 9) = ""
                        Vector(LugarVector, 10) = ""
                
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
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.Terminado = " + "'" + WTerminado + "'"
        ZSql = ZSql + " and Guia.Saldo <> 0"
        spMovguia = ZSql
        Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
        If rstMovguia.RecordCount > 0 Then
    
            With rstMovguia
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        LugarVector = LugarVector + 1
                        Vector(LugarVector, 1) = rstMovguia!Terminado
                        Vector(LugarVector, 2) = ""
                        Vector(LugarVector, 3) = rstMovguia!Fecha
                        Vector(LugarVector, 4) = Str$(rstMovguia!Lote)
                        Vector(LugarVector, 5) = Str$(rstMovguia!Saldo)
                        Vector(LugarVector, 6) = ""
                        Vector(LugarVector, 7) = ""
                        Vector(LugarVector, 8) = rstMovguia!Fecha
                        Vector(LugarVector, 9) = ""
                        Vector(LugarVector, 10) = ""
                    
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstMovguia.Close
        End If
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Entdev"
        ZSql = ZSql + " Where Entdev.Terminado = " + "'" + WTerminado + "'"
        ZSql = ZSql + " and Entdev.Saldo <> 0"
        spEntdev = ZSql
        Set rstEntdev = db.OpenRecordset(spEntdev, dbOpenSnapshot, dbSQLPassThrough)
        If rstEntdev.RecordCount > 0 Then
            With rstEntdev
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        LugarVector = LugarVector + 1
                        Vector(LugarVector, 1) = rstEntdev!Terminado
                        Vector(LugarVector, 2) = ""
                        Vector(LugarVector, 3) = rstEntdev!Fecha
                        Vector(LugarVector, 4) = Str$(rstEntdev!Lote)
                        Vector(LugarVector, 5) = Str$(rstEntdev!Saldo)
                        Vector(LugarVector, 6) = ""
                        Vector(LugarVector, 7) = ""
                        Vector(LugarVector, 8) = rstEntdev!Fecha
                        Vector(LugarVector, 9) = ""
                        Vector(LugarVector, 10) = ""
                
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstEntdev.Close
        End If
        
    Next Ciclo
    
    For Ciclo = 1 To LugarVector
   
        WTerminado = Vector(Ciclo, 1)
        WFecha = Vector(Ciclo, 3)
        WLote = Val(Vector(Ciclo, 4))
        WSaldo = Val(Vector(Ciclo, 5))
        
        ZSql = ""
        ZSql = ZSql + "Select * "
        ZSql = ZSql + " FROM Estadistica "
        ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WTerminado + "'"
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
                                
                        WWFecha = rstEstadistica!Fecha
                        WWNumero = rstEstadistica!Numero
                        
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
                        
                        WLoteAdicional = IIf(IsNull(rstEstadistica!LoteAdicional), "", rstEstadistica!LoteAdicional)
                        
                        If Len(Trim(WLoteAdicional)) = 98 Then
                            XLote6 = Mid$(WLoteAdicional, 1, 8)
                            XCanti6 = Mid$(WLoteAdicional, 9, 6)
                            XLote7 = Mid$(WLoteAdicional, 15, 8)
                            XCanti7 = Mid$(WLoteAdicional, 23, 6)
                            XLote8 = Mid$(WLoteAdicional, 29, 8)
                            XCanti8 = Mid$(WLoteAdicional, 37, 6)
                            XLote9 = Mid$(WLoteAdicional, 43, 8)
                            XCanti9 = Mid$(WLoteAdicional, 51, 6)
                            XLote10 = Mid$(WLoteAdicional, 57, 8)
                            XCanti10 = Mid$(WLoteAdicional, 65, 6)
                            XLote11 = Mid$(WLoteAdicional, 71, 8)
                            XCanti11 = Mid$(WLoteAdicional, 79, 6)
                            XLote12 = Mid$(WLoteAdicional, 85, 8)
                            XCanti12 = Mid$(WLoteAdicional, 93, 6)
                        End If
                        
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
                        XLote(6, 1) = Str$(ZLote6)
                        XLote(6, 2) = Str$(ZCanti6)
                        XLote(7, 1) = Str$(ZLote7)
                        XLote(7, 2) = Str$(ZCanti7)
                        XLote(8, 1) = Str$(ZLote8)
                        XLote(8, 2) = Str$(ZCanti8)
                        XLote(9, 1) = Str$(ZLote9)
                        XLote(9, 2) = Str$(ZCanti9)
                        XLote(10, 1) = Str$(ZLote10)
                        XLote(10, 2) = Str$(ZCanti10)
                        XLote(11, 1) = Str$(ZLote11)
                        XLote(11, 2) = Str$(ZCanti11)
                        XLote(12, 1) = Str$(ZLote12)
                        XLote(12, 2) = Str$(ZCanti12)
                    
                        If XLote(1, 2) = 0 Then
                            XLote(1, 2) = Str$(rstEstadistica!Cantidad)
                        End If
                        For x = 1 To 12
                            If Val(XLote(x, 1)) = WLote Then
                                WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                If WFecha2 > WFecha1 Then
                                    WFecha = WWFecha
                                    Vector(Ciclo, 6) = "Factura"
                                    Vector(Ciclo, 7) = Str$(WWNumero)
                                    Vector(Ciclo, 8) = WWFecha
                                    Vector(Ciclo, 9) = XLote(x, 2)
                                End If
                            End If
                        Next x
                
                        .MoveNext
                        If .EOF = True Then
                            Exit Do
                        End If
                
                    Loop
                End If
            End With
            rstEstadistica.Close
        End If
        
        
        
        ZSql = ""
        ZSql = ZSql + "Select * "
        ZSql = ZSql + " FROM Hoja "
        ZSql = ZSql + " Where Hoja.Terminado = " + "'" + WTerminado + "'"
        spHoja = ZSql
        Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
        If rstHoja.RecordCount > 0 Then
            With rstHoja
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                                
                        WWFecha = rstHoja!Fecha
                        WWNumero = rstHoja!Hoja
                        
                        Erase XLote
                
                        ZLote1 = IIf(IsNull(rstHoja!lote1), "0", rstHoja!lote1)
                        ZCanti1 = IIf(IsNull(rstHoja!Canti1), "0", rstHoja!Canti1)
                        ZLote2 = IIf(IsNull(rstHoja!lote2), "0", rstHoja!lote2)
                        ZCanti2 = IIf(IsNull(rstHoja!Canti2), "0", rstHoja!Canti2)
                        ZLote3 = IIf(IsNull(rstHoja!lote3), "0", rstHoja!lote3)
                        ZCanti3 = IIf(IsNull(rstHoja!Canti3), "0", rstHoja!Canti3)
                        
                        XLote(1, 1) = Str$(ZLote1)
                        XLote(1, 2) = Str$(ZCanti1)
                        XLote(2, 1) = Str$(ZLote2)
                        XLote(2, 2) = Str$(ZCanti2)
                        XLote(3, 1) = Str$(ZLote3)
                        XLote(3, 2) = Str$(ZCanti3)
                        
                        If Val(XLote(1, 1)) = 0 Then
                            XLote(1, 1) = Str$(rstHoja!Lote)
                            XLote(1, 2) = Str$(rstHoja!Cantidad)
                        End If
                        
                        For x = 1 To 3
                            If Val(XLote(x, 1)) = WLote Then
                                WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                                If WFecha2 > WFecha1 Then
                                    WFecha = WWFecha
                                    Vector(Ciclo, 6) = "Hoja"
                                    Vector(Ciclo, 7) = Str$(WWNumero)
                                    Vector(Ciclo, 8) = WWFecha
                                    Vector(Ciclo, 9) = XLote(x, 2)
                                End If
                            End If
                        Next x
                
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
        ZSql = ZSql + "Select * "
        ZSql = ZSql + " FROM Movvar "
        ZSql = ZSql + " Where Movvar.Terminado = " + "'" + WTerminado + "'"
        ZSql = ZSql + " and Movvar.Lote = " + "'" + Str$(WLote) + " '"
        ZSql = ZSql + " and Movvar.Movi = " + "'" + "S" + " '"
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
                        
                        WWCantidad = rstMovvar!Cantidad
                        WWFecha = rstMovvar!Fecha
                        WWNumero = rstMovvar!Codigo
                        WWLote = rstMovvar!Lote
                        
                        WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                        If WFecha2 > WFecha1 Then
                            WFecha = WWFecha
                            Vector(Ciclo, 6) = "Mov.Var."
                            Vector(Ciclo, 7) = Str$(WWNumero)
                            Vector(Ciclo, 8) = WWFecha
                            Vector(Ciclo, 9) = Str$(WWCantidad)
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
                        
                        If rstMovguia!Partida = WLote And rstMovguia!Movi = "S" Then
                        
                            WWCantidad = rstMovguia!Cantidad
                            WWFecha = rstMovguia!Fecha
                            WWNumero = rstMovguia!Codigo
                        
                            WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                            If WFecha2 > WFecha1 Then
                                WFecha = WWFecha
                                Vector(Ciclo, 6) = "Guia"
                                Vector(Ciclo, 7) = Str$(WWNumero)
                                Vector(Ciclo, 8) = WWFecha
                                Vector(Ciclo, 9) = Str$(WWCantidad)
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
        
        WDias = DateDiff("d", Vector(Ciclo, 8), Fecha.Text)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO HistoriaTerminado ("
        ZSql = ZSql + "Terminado ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Saldo ,"
        ZSql = ZSql + "Tipo ,"
        ZSql = ZSql + "Numero ,"
        ZSql = ZSql + "FechaII ,"
        ZSql = ZSql + "Cantidad ,"
        ZSql = ZSql + "Dias ,"
        ZSql = ZSql + "TituloI ,"
        ZSql = ZSql + "TituloII )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Vector(Ciclo, 1) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 3) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 4) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 5) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 6) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 7) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 8) + "',"
        ZSql = ZSql + "'" + Vector(Ciclo, 9) + "',"
        ZSql = ZSql + "'" + Str$(WDias) + "',"
        ZSql = ZSql + "'" + WTituloI + "',"
        ZSql = ZSql + "'" + Fecha.Text + "')"
        
        spHistoriaTerminado = ZSql
        Set rstHistoriaTerminado = db.OpenRecordset(spHistoriaTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    
    Next Ciclo
            
    
    Listado.WindowTitle = "Verificacion de Ultimos Movimientos de P.T."
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT HistoriaTerminado.Terminado, HistoriaTerminado.Fecha, HistoriaTerminado.Lote, HistoriaTerminado.Saldo, HistoriaTerminado.Tipo, HistoriaTerminado.Numero, HistoriaTerminado.FechaII, HistoriaTerminado.Cantidad, HistoriaTerminado.Dias, HistoriaTerminado.TituloI, HistoriaTerminado.TituloII, " _
                    + "Terminado.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.HistoriaTerminado HistoriaTerminado, " _
                    + DSQ + ".dbo.Terminado Terminado " _
                    + "Where " _
                    + "HistoriaTerminado.Terminado = Terminado.Codigo AND " _
                    + "HistoriaTerminado.Terminado >= '" + Desde.Text + "' AND " _
                    + "HistoriaTerminado.Terminado <= '" + Hasta.Text + "'"
                    
    Listado.Connect = Connect()
    Listado.GroupSelectionFormula = "{HistoriaTerminado.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    
    Exit Sub
    
WError:
     Resume Next
    
End Sub


Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    DbsEmpresa.Close
    
    PrgHistoriaTerminado.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Desde.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Fecha.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_FichaTer
End Sub


Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgHistoriaTerminado.Caption = "Listado de Verificacion de Ultimos Movimientos de P.T. :  " + !Nombre
        End If
    End With
    Fecha.Text = "  /  /    "
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spTerminado = "ListaTerminado"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstTerminado
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstTerminado!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstTerminado.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spTerminado = "ConsultaTerminado " + "'" + Claveven$ + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        Desde.Text = rstTerminado!Codigo
        Hasta.Text = rstTerminado!Codigo
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


