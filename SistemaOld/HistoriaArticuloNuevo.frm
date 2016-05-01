VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgHistoriaArticulo 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Verificacion de Ultimos Movimientos de M.P."
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
         Mask            =   "AA-###-###"
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
         Mask            =   "AA-###-###"
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
      ReportFileName  =   "WHistoriaArticulo.rpt"
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
      ItemData        =   "HistoriaArticuloNuevo.frx":0000
      Left            =   120
      List            =   "HistoriaArticuloNuevo.frx":0007
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
Attribute VB_Name = "PrgHistoriaArticulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WEntrada As Double
Private WSalida As Double
Private WTipo As Integer
Private WNumero As String
Private Impre1 As String
Private Impre2 As String
Private WFecha As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovvar As Recordset
Dim spMovguia As String
Dim rstMovguia As Recordset
Dim spMovvar As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstHistoriaTerminado As Recordset
Dim spHistoriaTerminado As String
Dim spGuia As String
Dim rstGuia As Recordset
Dim XParam As String
Dim Vector(10000, 10) As String
Dim ZArticulo(10000) As String
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

Dim ZDias As String
Dim ZComparaI As Date
Dim ZComparaII As Date
Dim ZStock As Double


Private Sub Acepta_Click()

    Rem On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    Erase ZArticulo
    LugarArticulo = 0
    
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
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Articulo.Codigo >= " + "'" + Desde.Text + "'"
    ZSql = ZSql + " and Articulo.Codigo <= " + "'" + Hasta.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        With rstArticulo
            .MoveFirst
            If .NoMatch = False Then
                Do
                    If .EOF = True Then
                        Exit Do
                    End If
                
                    WArticulo = rstArticulo!Codigo
                    ZStock = rstArticulo!Entradas - rstArticulo!Salidas
                    Call Redondeo(ZStock)
                    If ZStock > 0 Then
                        LugarArticulo = LugarArticulo + 1
                        ZArticulo(LugarArticulo) = WArticulo
                    End If
                    
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
        rstArticulo.Close
    End If
    
    Erase Vector
    LugarVector = 0
    
    
    For Ciclo = 1 To LugarArticulo
        
        WArticulo = ZArticulo(Ciclo)
        
        ZSql = ""
        ZSql = ZSql + "Select Laudo.Articulo, Laudo.Saldo, Laudo.Fecha, Laudo.Laudo, Laudo.Saldo"
        ZSql = ZSql + " FROM Laudo"
        ZSql = ZSql + " Where Laudo.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and Laudo.Saldo <> 0"
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
    
            With rstLaudo
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        ZStock = rstLaudo!Saldo
                        Call Redondeo(ZStock)
                        If ZStock > 0 Then
                            LugarVector = LugarVector + 1
                            Vector(LugarVector, 1) = rstLaudo!Articulo
                            Vector(LugarVector, 2) = ""
                            Vector(LugarVector, 3) = rstLaudo!Fecha
                            Vector(LugarVector, 4) = Str$(rstLaudo!Laudo)
                            Vector(LugarVector, 5) = Str$(rstLaudo!Saldo)
                            Vector(LugarVector, 6) = ""
                            Vector(LugarVector, 7) = ""
                            Vector(LugarVector, 8) = rstLaudo!Fecha
                            Vector(LugarVector, 9) = ""
                            Vector(LugarVector, 10) = ""
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
        
        
        
    
        ZSql = ""
        ZSql = ZSql + "Select Guia.Articulo, Guia.Saldo, Guia.Fecha, Guia.Lote, Guia.Saldo"
        ZSql = ZSql + " FROM Guia"
        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and Guia.Saldo <> 0"
        spGuia = ZSql
        Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
        If rstGuia.RecordCount > 0 Then
    
            With rstGuia
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                
                        ZStock = rstGuia!Saldo
                        Call Redondeo(ZStock)
                        If ZStock > 0 Then
                            LugarVector = LugarVector + 1
                            Vector(LugarVector, 1) = rstGuia!Articulo
                            Vector(LugarVector, 2) = ""
                            Vector(LugarVector, 3) = rstGuia!Fecha
                            Vector(LugarVector, 4) = Str$(rstGuia!Lote)
                            Vector(LugarVector, 5) = Str$(rstGuia!Saldo)
                            Vector(LugarVector, 6) = ""
                            Vector(LugarVector, 7) = ""
                            Vector(LugarVector, 8) = rstGuia!Fecha
                            Vector(LugarVector, 9) = ""
                            Vector(LugarVector, 10) = ""
                        End If
                    
                        .MoveNext
                
                        If .EOF = True Then
                            Exit Do
                        End If
                    Loop
                End If
            End With
            rstGuia.Close
        End If
        
    Next Ciclo
    
    For Ciclo = 1 To LugarVector
   
        WArticulo = Vector(Ciclo, 1)
        WFecha = Vector(Ciclo, 3)
        WLote = Val(Vector(Ciclo, 4))
        WSaldo = Val(Vector(Ciclo, 5))
        
        WArticuloDy = Left$(WArticulo, 3) + "00" + Right$(WArticulo, 7)
        
        ZSql = ""
        ZSql = ZSql + "Select Estadistica.Articulo, Estadistica.OrdFecha, Estadistica.fecha, Estadistica.Numero, Estadistica.Lote1, Estadistica.Canti1, Estadistica.Lote2, Estadistica.Canti2, Estadistica.Lote3, Estadistica.Canti3, Estadistica.Lote4, Estadistica.Canti4, Estadistica.Lote5, Estadistica.Canti5, Estadistica.cantidad, Estadistica.LoteAdicional "
        ZSql = ZSql + " FROM Estadistica "
        ZSql = ZSql + " Where Estadistica.Articulo = " + "'" + WArticuloDy + "'"
        ZSql = ZSql + " order by Estadistica.OrdFecha desc"
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
                    
                        If Val(XLote(1, 2)) = 0 Then
                            XLote(1, 2) = rstEstadistica!Cantidad
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
                                Exit Do
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
        ZSql = ZSql + "Select Hoja.Articulo, Hoja.FechaOrd, Hoja.Fecha, Hoja.Hoja, Hoja.Lote1, Hoja.Canti1, Hoja.Lote2, Hoja.Canti2, Hoja.Lote3, Hoja.Canti3, Hoja.Lote, Hoja.Cantidad "
        ZSql = ZSql + " FROM Hoja "
        ZSql = ZSql + " Where Hoja.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " Order by Hoja.FechaOrd desc"
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
                                Exit Do
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
        ZSql = ZSql + " Where Movvar.Articulo = " + "'" + WArticulo + "'"
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
    
    
    
        ZSql = ""
        ZSql = ZSql + "Select * "
        ZSql = ZSql + " FROM Guia "
        ZSql = ZSql + " Where Guia.Articulo = " + "'" + WArticulo + "'"
        ZSql = ZSql + " and Guia.Lote = " + "'" + Str$(WLote) + " '"
        ZSql = ZSql + " and Guia.Movi = " + "'" + "S" + " '"
        spGuia = ZSql
        Set rstGuia = db.OpenRecordset(spGuia, dbOpenSnapshot, dbSQLPassThrough)
        If rstGuia.RecordCount > 0 Then
            With rstGuia
                .MoveFirst
                If .NoMatch = False Then
                    Do
                        If .EOF = True Then
                            Exit Do
                        End If
                        
                        WWCantidad = rstGuia!Cantidad
                        WWFecha = rstGuia!Fecha
                        WWNumero = rstGuia!Codigo
                        WWLote = rstGuia!Lote
                        
                        WFecha1 = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        WFecha2 = Right$(WWFecha, 4) + Mid$(WWFecha, 4, 2) + Left$(WWFecha, 2)
                        If WFecha2 > WFecha1 Then
                            WFecha = WWFecha
                            Vector(Ciclo, 6) = "Guia"
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
            rstGuia.Close
        End If
        
        ZComparaI = "01/01/1900"
        ZComparaII = "01/01/1900"
        
        ZComparaI = Vector(Ciclo, 8)
        ZComparaII = Fecha.Text
        
        WDias = DateDiff("d", ZComparaI, ZComparaII)
    
        ZSql = ""
        ZSql = ZSql + "INSERT INTO HistoriaTerminado ("
        ZSql = ZSql + "Articulo ,"
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
            
        
        
    
    Listado.WindowTitle = "Verificacion de Ultimos Movimientos de M.P."
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT HistoriaTerminado.Fecha, HistoriaTerminado.Lote, HistoriaTerminado.Saldo, HistoriaTerminado.Tipo, HistoriaTerminado.Numero, HistoriaTerminado.FechaII, HistoriaTerminado.Cantidad, HistoriaTerminado.Dias, HistoriaTerminado.TituloI, HistoriaTerminado.TituloII, HistoriaTerminado.Articulo, " _
                + "Articulo.Descripcion " _
                + "From " _
                + DSQ + ".dbo.HistoriaTerminado HistoriaTerminado, " _
                + DSQ + ".dbo.Articulo Articulo " _
                + "Where " _
                + "HistoriaTerminado.Articulo = Articulo.Codigo AND " _
                + "HistoriaTerminado.Articulo >= ' ' AND " _
                + "HistoriaTerminado.Articulo <= 'ZZZZZZZZZZ'"
                    
    Listado.Connect = Connect()
    Rem Listado.GroupSelectionFormula = "{HistoriaTerminado.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    
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
    
    PrgHistoriaArticulo.Hide
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
            PrgHistoriaArticulo.Caption = "Listado de Verificacion de Ultimos Movimientos de M.P. :  " + !Nombre
        End If
    End With
    Fecha.Text = "  /  /    "
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spArticulo = "ListaArticulo"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    With rstArticulo
        .MoveFirst
            Do
            If .EOF = False Then
                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                Pantalla.AddItem IngresaItem
                IngresaItem = rstArticulo!Codigo
                WIndice.AddItem IngresaItem
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    rstArticulo.Close
            
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    
    Indice = Pantalla.ListIndex
    Claveven$ = WIndice.List(Indice)
    spArticulo = "ConsultaArticulo " + "'" + Claveven$ + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Desde.Text = rstArticulo!Codigo
        Hasta.Text = rstArticulo!Codigo
            Else
        Desde.Text = Claveven$
        Hasta.Text = Claveven$
    End If
    Desde.SetFocus
    
End Sub


