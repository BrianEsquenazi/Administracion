VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAnalisisPedido 
   AutoRedraw      =   -1  'True
   Caption         =   "Analisis de Cumplimiento de Pedidos de Venta"
   ClientHeight    =   4575
   ClientLeft      =   2025
   ClientTop       =   1050
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   4575
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   4095
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   7335
      Begin VB.ComboBox TipoFechaII 
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
         TabIndex        =   14
         Top             =   2640
         Width           =   4815
      End
      Begin VB.ComboBox TipoFecha 
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
         TabIndex        =   12
         Top             =   2160
         Width           =   4815
      End
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
         TabIndex        =   11
         Top             =   1680
         Width           =   4815
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2280
         TabIndex        =   9
         Top             =   1200
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
         Top             =   720
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
         Left            =   3840
         TabIndex        =   8
         Top             =   3240
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
         Left            =   2040
         TabIndex        =   7
         Top             =   3240
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Entrega"
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
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Pedido"
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
         TabIndex        =   13
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo"
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
         TabIndex        =   10
         Top             =   1680
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
         TabIndex        =   4
         Top             =   1200
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
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlistinf.rpt"
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
      Left            =   6600
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgAnalisisPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WVector(20000, 11) As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim WSolicitud As Integer
Dim WConcepto As Integer
Dim WProblema As String
Dim WMp As String
Dim WFechaSolicitud As String
Dim ZProblema(10, 4) As String
Dim Ciclo As Integer

Dim Posi As Integer
Dim XFec1 As String
Dim XFec2 As String
Dim SumaDia As Integer
Dim DiaFeriado(100) As String
Private TotalFeriado As Integer

Private Sub Acepta_Click()

    On Error GoTo WError
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            WDesEmpresa = !Nombre
        End If
    End With
    
    Erase DiaFeriado
    TotalFeriado = 0
    
    spFeriado = "ListaFeriado"
    Set rstFeriado = db.OpenRecordset(spFeriado, dbOpenSnapshot, dbSQLPassThrough)
    If rstFeriado.RecordCount > 0 Then
        With rstFeriado
            .MoveFirst
            Do
                If .EOF = False Then
                    TotalFeriado = TotalFeriado + 1
                    DiaFeriado(TotalFeriado) = rstFeriado!Fecha
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFeriado.Close
    End If
    
    
    WDesde = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    WHasta = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    WTitulo = "" + Desde.Text + " al " + Hasta.Text
    If Tipo.ListIndex <> 4 And Tipo.ListIndex <> 5 Then
        If TipoFecha.ListIndex = 0 Then
            WTitulo = WTitulo + " (F.Pactada)"
                Else
            WTitulo = WTitulo + " (F.Original)"
        End If
    End If
    If TipoFechaII.ListIndex = 0 Then
        WTitulo = WTitulo + "-(F.Factura)"
            Else
        WTitulo = WTitulo + "-(F.Entrega)"
    End If
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Pedido SET "
    ZSql = ZSql + " Suma1 = 0,"
    ZSql = ZSql + " Suma2 = 0,"
    ZSql = ZSql + " Dias = 0,"
    ZSql = ZSql + " TipoFecha = " + " '" + Str$(TipoFechaII.ListIndex) + "',"
    ZSql = ZSql + " Titulo = " + " '" + WTitulo + "'"
    spPedido = ZSql
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    
    Erase WVector
    Renglon = 0
    
    Sql1 = "Select Pedido, Fecha, FechaOrd, Terminado, Clave, FecEntrega, OrdFecEntrega, Cliente, FechaInicial, OrdFechaInicial, TipoPed, FechaActualizacion, OrdFechaActualizacion"
    Sql2 = " FROM Pedido"
    Sql3 = " Where Pedido.FechaOrd >= " + "'" + WDesde + "'"
    Sql4 = " and Pedido.FechaOrd <= " + "'" + WHasta + "'"
    Sql5 = " and Pedido.TipoPed <> 5"
    spPedido = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
    If rstPedido.RecordCount > 0 Then
        With rstPedido
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    
                    WVector(Renglon, 1) = rstPedido!Pedido
                    WVector(Renglon, 2) = rstPedido!Fecha
                    WVector(Renglon, 4) = rstPedido!Terminado
                    WVector(Renglon, 5) = rstPedido!Clave
                    WVector(Renglon, 6) = rstPedido!FechaOrd
                    If TipoFecha.ListIndex = 0 Then
                        WVector(Renglon, 3) = rstPedido!FecEntrega
                        WVector(Renglon, 7) = rstPedido!OrdFecEntrega
                            Else
                        XFechaInicial = IIf(IsNull(rstPedido!FechaInicial), "", rstPedido!FechaInicial)
                        XOrdFechaInicial = IIf(IsNull(rstPedido!OrdFechaInicial), "", rstPedido!OrdFechaInicial)
                        If XFechaInicial <> "" Then
                            WVector(Renglon, 3) = rstPedido!FechaInicial
                            WVector(Renglon, 7) = rstPedido!OrdFechaInicial
                                Else
                            WVector(Renglon, 3) = rstPedido!FecEntrega
                            WVector(Renglon, 7) = rstPedido!OrdFecEntrega
                        End If
                    End If
                    WVector(Renglon, 8) = rstPedido!Cliente
                    WVector(Renglon, 9) = rstPedido!Tipoped
                    XFechaActualizacion = IIf(IsNull(rstPedido!FechaActualizacion), "", rstPedido!FechaActualizacion)
                    XOrdFechaActualizacion = IIf(IsNull(rstPedido!OrdFechaActualizacion), "", rstPedido!OrdFechaActualizacion)
                    WVector(Renglon, 10) = XFechaActualizacion
                    WVector(Renglon, 11) = XOrdFechaActualizacion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPedido.Close
    End If
    
  For Ciclo = 1 To Renglon
    
        WPedido = WVector(Ciclo, 1)
        WFecha = WVector(Ciclo, 2)
        WFechaEntrega = WVector(Ciclo, 3)
        WTerminado = WVector(Ciclo, 4)
        WClave = WVector(Ciclo, 5)
        WFechaord = WVector(Ciclo, 6)
        WOrdFechaEntrega = WVector(Ciclo, 7)
        WCliente = WVector(Ciclo, 8)
        WTipoPedido = WVector(Ciclo, 9)
        WFechaActualizacion = WVector(Ciclo, 10)
        WOrdFechaActualizacion = WVector(Ciclo, 11)
        
        If Val(WTipoPedido) = 4 Then
        
            If WFechaActualizacion <> "" Then
                Entra = "S"
                FechaFactu = WFechaActualizacion
                FechaFactuOrd = WOrdFechaActualizacion
                    Else
                Entra = "N"
                FechaFactu = ""
                FechaFactuOrd = ""
            End If
                    
                Else
        
            Entra = "N"
            FechaFactu = ""
            FechaFactuOrd = ""
            ZZRemito = ""
        
            Sql1 = "Select Fecha, OrdFecha, Articulo, Pedido, Remito "
            Sql2 = " FROM Estadistica"
            Sql3 = " Where Estadistica.Pedido = " + "'" + WPedido + "'"
            Sql4 = " and Estadistica.Articulo = " + "'" + WTerminado + "'"
            Sql5 = " order by OrdFecha"
            spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEstadistica.RecordCount > 0 Then
                With rstEstadistica
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Entra = "S"
                            FechaFactu = rstEstadistica!Fecha
                            FechaFactuOrd = rstEstadistica!OrdFecha
                            ZZRemito = rstEstadistica!Remito
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEstadistica.Close
            End If
            
            If TipoFechaII.ListIndex = 1 Then
                If ZZRemito <> "" Then
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM HojaRuta"
                    ZSql = ZSql + " Where HojaRuta.Pedido = " + "'" + WPedido + "'"
                    ZSql = ZSql + " and HojaRuta.Remito = " + "'" + Trim(ZZRemito) + "'"
                    spHojaRuta = ZSql
                    Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHojaRuta.RecordCount > 0 Then
                        FechaFactu = rstHojaRuta!Fecha
                        FechaFactuOrd = rstHojaRuta!OrdFecha
                        rstHojaRuta.Close
                            Else
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM HojaRuta"
                        ZSql = ZSql + " Where HojaRuta.Pedido = " + "'" + WPedido + "'"
                        spHojaRuta = ZSql
                        Set rstHojaRuta = db.OpenRecordset(spHojaRuta, dbOpenSnapshot, dbSQLPassThrough)
                        If rstHojaRuta.RecordCount > 0 Then
                            FechaFactu = rstHojaRuta!Fecha
                            FechaFactuOrd = rstHojaRuta!OrdFecha
                            rstHojaRuta.Close
                        End If
                    End If
                End If
            End If
            
        End If
        
        If Entra = "S" Then
                    
            If Left$(WTerminado, 2) = "PT" Or Left$(WTerminado, 2) = "PE" Then
            
                WLinea = ""
                Sql1 = "Select *"
                Sql2 = " FROM Terminado"
                Sql3 = " Where Terminado.Codigo = " + "'" + WTerminado + "'"
                spTerminado = Sql1 + Sql2 + Sql3
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WLinea = Str$(rstTerminado!Linea)
                    rstTerminado.Close
                End If
                
                    Else
                    
                If Left$(WTerminado, 2) = "DY" Then
                    WLinea = "16"
                        Else
                    If Left$(WTerminado, 2) = "DS" Then
                        WLinea = "16"
                            Else
                        If Left$(WTerminado, 2) = "DQ" Then
                            WLinea = "22"
                                Else
                            If Left$(WTerminado, 2) = "DW" Then
                                WLinea = "17"
                                    Else
                                WLinea = "5"
                            End If
                        End If
                    End If
                End If
                
            End If
        
            WDias = 0
            WSuma2 = "0"
            
            WFechaDesdeOrd = WOrdFechaEntrega
            WFechaDesde = WFechaEntrega
            If Tipo.ListIndex = 4 Then
                WFechaDesdeOrd = WFechaord
                WFechaDesde = WFecha
            End If
            If Tipo.ListIndex = 5 Then
                WFechaDesdeOrd = WFechaord
                WFechaDesde = WFecha
            End If
            
            WFechaHastaOrd = FechaFactuOrd
            WFechaHasta = FechaFactu
            
            
            If WFechaHastaOrd > WFechaDesdeOrd Then
            
                WSuma2 = "1"
                Rem by nan
                Do
        
                    Feriado = "N"
                    For Cicla = 1 To TotalFeriado
                        If DiaFeriado(Cicla) = WFechaDesde Then
                            Feriado = "S"
                            Exit For
                        End If
                    Next Cicla
                    
                    Rem 1 - DOMINGO
                    Rem 2 - LUNES
                    Rem 3 - MARTES
                    Rem 4 - MIERCOLES
                    Rem 5 - JUEVES
                    Rem 6 - VIERNES
                    Rem 7 - SABADO
                    
                    XFec1 = WFechaDesde
                    strDia = Format$(XFec1, "dddd")
                    BDia = Format(XFec1, "w")
                    If BDia = 1 Or BDia = 7 Then
                        Feriado = "S"
                    End If
            
                    If Feriado = "N" Then
                        WDias = WDias + 1
                    End If
                    SumaDia = 2
                    Call Calcula_vencimiento(XFec1, SumaDia, XFec2)
                    WFechaDesde = XFec2
                        
                    If WFechaDesde = WFechaHasta Then
                        Exit Do
                    End If
        
                Loop
        
            End If
            
            DifeDias = Str$(WDias)
            WSuma1 = "1"
            If WDias <= 0 Then
                WSuma2 = 0
            End If
            
            Select Case Val(WLinea)
                Case 3, 4, 5, 7, 8, 11, 12, 14
                    WSumaLinea = "1"
                    WDesSumaLinea = "QUIMICOS"
                Case 6, 16, 17
                    WSumaLinea = "2"
                    WDesSumaLinea = "COLORANTES"
                Case 10, 22, 24, 25, 26, 27, 28, 29, 30
                    WSumaLinea = "3"
                    WDesSumaLinea = "FARMA"
                Case 20
                    WSumaLinea = "5"
                    WDesSumaLinea = "FAZONES FARMA"
                Case 21
                    WSumaLinea = "6"
                    WDesSumaLinea = "FAZONES QUIMICOS"
                Case Else
                    WRubro = 0
                    spCliente = "ConsultaCliente " + WCliente
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        WRubro = rstCliente!Rubro
                        rstCliente.Close
                    End If
                    If WCliente = "P00005" Then
                        WSumaLinea = "4"
                        WDesSumaLinea = "FAZONES PELLITAL"
                            Else
                        If WRubro = 10 Then
                            WSumaLinea = "5"
                            WDesSumaLinea = "FAZONES FARMA"
                                Else
                            WSumaLinea = "6"
                            WDesSumaLinea = "FAZONES QUIMICOS"
                        End If
                    End If
            End Select
            
            If Tipo.ListIndex = 4 Or Tipo.ListIndex = 5 Then
                If Val(WTipoPedido) = 0 Then
                    ZProvincia = 0
                    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
                    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                    If rstCliente.RecordCount > 0 Then
                        ZProvincia = rstCliente!Provincia
                        rstCliente.Close
                    End If
                    If ZProvincia <> 24 Then
                        WSumaLinea = "1"
                        WDesSumaLinea = "Normal"
                            Else
                        WSumaLinea = "2"
                        WDesSumaLinea = "Normal Expo"
                    End If
                        Else
                    WSumaLinea = "3"
                    WDesSumaLinea = "Resto"
                End If
            End If
            
            WSolicitud = 0
            WMp = ""
            WFechaSolicitud = ""
            Erase ZProblema
            ZLugar = 0
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Atraso"
            ZSql = ZSql + " Where Atraso.Pedido = " + "'" + WPedido + "'"
            ZSql = ZSql + " Order by Atraso.Numero"
            spAtraso = ZSql
            Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
            If rstAtraso.RecordCount > 0 Then
                With rstAtraso
                    .MoveFirst
                    If .NoMatch = False Then
                        Do
                            ZOrigen = IIf(IsNull(!Origen), "0", !Origen)
                            Select Case ZOrigen
                                Case 0
                                    WConcepto = IIf(IsNull(rstAtraso!Concepto), "0", rstAtraso!Concepto)
                                    WProblema = IIf(IsNull(rstAtraso!Problema), "", rstAtraso!Problema)
                                    WSolicitud = IIf(IsNull(rstAtraso!Solicitud), "0", rstAtraso!Solicitud)
                                    WMp = IIf(IsNull(rstAtraso!Articulo), "", rstAtraso!Articulo)
                                    ZLugar = ZLugar + 1
                                    ZProblema(ZLugar, 1) = WConcepto
                                    ZProblema(ZLugar, 2) = WProblema
                                    ZProblema(ZLugar, 3) = ZOrigen
                                    ZProblema(ZLugar, 4) = Left$(rstAtraso!Fecha, 5)
                                Case Else
                                    WConcepto = IIf(IsNull(rstAtraso!Concepto), "0", rstAtraso!Concepto)
                                    WProblema = IIf(IsNull(rstAtraso!Problema), "", rstAtraso!Problema)
                                    ZLugar = ZLugar + 1
                                    ZProblema(ZLugar, 1) = WConcepto
                                    ZProblema(ZLugar, 2) = WProblema
                                    ZProblema(ZLugar, 3) = ZOrigen
                                    ZProblema(ZLugar, 4) = Left$(rstAtraso!Fecha, 5)
                            End Select
                            .MoveNext
                            If .EOF = True Then
                                Exit Do
                            End If
                        Loop
                    End If
                End With
                rstAtraso.Close
            End If
            
            If WSolicitud <> 0 Then
                Sql1 = "Select *"
                Sql2 = " FROM Solic"
                Sql3 = " Where Solic.Solicitud = " + "'" + Str$(WSolicitud) + "'"
                spSolic = Sql1 + Sql2 + Sql3
                Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                If rstSolic.RecordCount > 0 Then
                    WFechaSolicitud = IIf(IsNull(rstSolic!Fecha), "0", rstSolic!Fecha)
                    rstSolic.Close
                End If
            End If
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Pedido SET "
            ZSql = ZSql + "Suma1 = " + "'" + WSuma1 + "',"
            ZSql = ZSql + "Suma2 = " + "'" + Str$(WSuma2) + "',"
            ZSql = ZSql + "Dias = " + "'" + DifeDias + "',"
            ZSql = ZSql + "FechaReal = " + "'" + FechaFactu + "',"
            ZSql = ZSql + "Linea = " + "'" + WLinea + "',"
            ZSql = ZSql + "SumaLinea = " + "'" + WSumaLinea + "',"
            ZSql = ZSql + "DesSumaLinea = " + "'" + WDesSumaLinea + "',"
            ZSql = ZSql + "DesEmpresa = " + "'" + WDesEmpresa + "',"
            ZSql = ZSql + "OrdFechaReal = " + "'" + FechaFactuOrd + "',"
            ZSql = ZSql + "Concepto = " + "'" + ZProblema(1, 1) + "',"
            ZSql = ZSql + "Problema = " + "'" + ZProblema(1, 2) + "',"
            ZSql = ZSql + "OrigenI = " + "'" + ZProblema(1, 3) + "',"
            ZSql = ZSql + "FechaI = " + "'" + ZProblema(1, 4) + "',"
            ZSql = ZSql + "ConceptoII = " + "'" + ZProblema(2, 1) + "',"
            ZSql = ZSql + "ProblemaII = " + "'" + ZProblema(2, 2) + "',"
            ZSql = ZSql + "OrigenII = " + "'" + ZProblema(2, 3) + "',"
            ZSql = ZSql + "FechaII = " + "'" + ZProblema(2, 4) + "',"
            ZSql = ZSql + "ConceptoIII = " + "'" + ZProblema(3, 1) + "',"
            ZSql = ZSql + "ProblemaIII = " + "'" + ZProblema(3, 2) + "',"
            ZSql = ZSql + "OrigenIII = " + "'" + ZProblema(3, 3) + "',"
            ZSql = ZSql + "FechaIII = " + "'" + ZProblema(3, 4) + "',"
            ZSql = ZSql + "ConceptoIV = " + "'" + ZProblema(4, 1) + "',"
            ZSql = ZSql + "ProblemaIV = " + "'" + ZProblema(4, 2) + "',"
            ZSql = ZSql + "OrigenIV = " + "'" + ZProblema(4, 3) + "',"
            ZSql = ZSql + "FechaIV = " + "'" + ZProblema(4, 4) + "',"
            ZSql = ZSql + "Mp = " + "'" + WMp + "',"
            ZSql = ZSql + "FechaSolicitud = " + "'" + WFechaSolicitud + "'"
            ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
                     
            spPedido = ZSql
            Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
            
       End If
   Rem by nan
   Next Ciclo
    
    Listado.WindowTitle = "Analisis de Cumplimiento de Pedidos de Venta"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Select Case Tipo.ListIndex
        Case 0
            Listado.GroupSelectionFormula = "{Pedido.Suma1} in 1 to 1"
            Listado.SelectionFormula = "{Pedido.Suma1} in 1 to 1"
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, " _
                            + "Cliente.Razon, " _
                            + "Lineas.Nombre " _
                            + "From " _
                            + DSQ + ".dbo.Pedido Pedido, " _
                            + DSQ + ".dbo.Cliente Cliente, " _
                            + DSQ + ".dbo.Lineas Lineas " _
                            + "Where " _
                            + "Pedido.Cliente = Cliente.Cliente AND " _
                            + "Pedido.Linea = Lineas.Linea AND " _
                            + "Pedido.Suma1 = 1"
                
            Listado.ReportFileName = "analisispedresu.rpt"
            
        Case 1
            Listado.GroupSelectionFormula = "{Pedido.Dias} in 1 to 999999 and {Pedido.Suma1} in 1 to 1"
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Version, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, Pedido.FechaInicial, Pedido.Concepto, Pedido.Problema, Pedido.Mp, Pedido.FechaSolicitud, Pedido.ConceptoII, Pedido.ProblemaII, Pedido.TipoFecha, Pedido.ConceptoIII, Pedido.ProblemaIII, Pedido.ConceptoIV, Pedido.ProblemaIV, Pedido.OrigenI, Pedido.OrigenII, Pedido.OrigenIII, Pedido.OrigenIV, Pedido.FechaI, Pedido.FechaII, Pedido.FechaIII, Pedido.FechaIV, " _
                        + "Cliente.Razon, " _
                        + "Lineas.Nombre " _
                        + "From " _
                        + DSQ + ".dbo.Pedido Pedido, " _
                        + DSQ + ".dbo.Cliente Cliente, " _
                        + DSQ + ".dbo.Lineas Lineas " _
                        + "Where " _
                        + "Pedido.Cliente = Cliente.Cliente AND " _
                        + "Pedido.Linea = Lineas.Linea AND " _
                        + "Pedido.Suma1 = 1 AND " _
                        + "Pedido.Dias >= 1 AND " _
                        + "Pedido.Dias <= 999999"
            
            Listado.ReportFileName = "analisisped.rpt"
            
        Case 2
            Listado.GroupSelectionFormula = "{Pedido.Dias} in -999999 to 999999 and {Pedido.Suma1} in 1 to 1"
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, Pedido.FechaInicial, Pedido.TipoFecha, Pedido.Version, " _
                            + "Cliente.Razon, " _
                            + "Lineas.Nombre " _
                            + "From " _
                            + DSQ + ".dbo.Pedido Pedido, " _
                            + DSQ + ".dbo.Cliente Cliente, " _
                            + DSQ + ".dbo.Lineas Lineas " _
                            + "Where " _
                            + "Pedido.Cliente = Cliente.Cliente AND " _
                            + "Pedido.Linea = Lineas.Linea AND " _
                            + "Pedido.Suma1 = 1 AND " _
                            + "Pedido.Dias >= -999999 AND " _
                            + "Pedido.Dias <= 999999"
                
            Listado.ReportFileName = "analisispedTotal.rpt"
            
        Case 3
            Listado.GroupSelectionFormula = "{Pedido.Dias} in 1 to 999999 and {Pedido.Suma1} in 1 to 1"
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, Pedido.FechaInicial, Pedido.Concepto, Pedido.Problema, Pedido.Mp, Pedido.FechaSolicitud, Pedido.TipoFecha, " _
                            + "Cliente.Razon, " _
                            + "Lineas.Nombre " _
                            + "From " _
                            + DSQ + ".dbo.Pedido Pedido, " _
                            + DSQ + ".dbo.Cliente Cliente, " _
                            + DSQ + ".dbo.Lineas Lineas " _
                            + "Where " _
                            + "Pedido.Cliente = Cliente.Cliente AND " _
                            + "Pedido.Linea = Lineas.Linea AND " _
                            + "Pedido.Suma1 = 1 AND " _
                            + "Pedido.Dias >= 1 AND " _
                            + "Pedido.Dias <= 999999"
            
            Listado.ReportFileName = "analisispedconcepto.rpt"
            
        Case 4
            Rem Listado.GroupSelectionFormula = "{Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Listado.GroupSelectionFormula = ""
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, " _
                            + "Cliente.Razon, " _
                            + "Lineas.Nombre " _
                            + "From " _
                            + DSQ + ".dbo.Pedido Pedido, " _
                            + DSQ + ".dbo.Cliente Cliente, " _
                            + DSQ + ".dbo.Lineas Lineas " _
                            + "Where " _
                            + "Pedido.Cliente = Cliente.Cliente AND " _
                            + "Pedido.Linea = Lineas.Linea AND " _
                            + "Pedido.Suma1 = 1"
                
            Listado.ReportFileName = "analisispedOtro.rpt"
            
        Case 5
            Rem Listado.GroupSelectionFormula = "{Informe.fechaord} in " + Chr$(34) + WDesde + Chr$(34) + " to " + Chr$(34) + WHasta + Chr$(34)
            Listado.GroupSelectionFormula = ""
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, " _
                            + "Cliente.Razon, " _
                            + "Lineas.Nombre " _
                            + "From " _
                            + DSQ + ".dbo.Pedido Pedido, " _
                            + DSQ + ".dbo.Cliente Cliente, " _
                            + DSQ + ".dbo.Lineas Lineas " _
                            + "Where " _
                            + "Pedido.Cliente = Cliente.Cliente AND " _
                            + "Pedido.Linea = Lineas.Linea AND " _
                            + "Pedido.Suma1 = 1"
                
            Listado.ReportFileName = "analisispedOtroResu.rpt"
            
        Case Else
            Listado.GroupSelectionFormula = "{Pedido.Dias} in 1 to 999999 and {Pedido.Suma1} in 1 to 1"
       
            DbConnect = db.Connect
            DSQ = getDatabase(DbConnect)
    
            Listado.SQLQuery = "SELECT Pedido.Pedido, Pedido.Cliente, Pedido.Fecha, Pedido.FecEntrega, Pedido.Terminado, Pedido.Cantidad, Pedido.Linea, Pedido.Suma1, Pedido.Suma2, Pedido.Dias, Pedido.FechaReal, Pedido.Titulo, Pedido.SumaLinea, Pedido.DesSumaLinea, Pedido.DesEmpresa, Pedido.FechaInicial, Pedido.Concepto, Pedido.Problema, Pedido.Mp, Pedido.FechaSolicitud, Pedido.TipoFecha, " _
                    + "Cliente.Razon, " _
                    + "Lineas.Nombre " _
                    + "From " _
                    + DSQ + ".dbo.Pedido Pedido, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Lineas Lineas " _
                    + "Where " _
                    + "Pedido.Cliente = Cliente.Cliente AND " _
                    + "Pedido.Linea = Lineas.Linea AND " _
                    + "Pedido.Suma1 = 1 AND " _
                    + "Pedido.Dias >= 1 AND " _
                    + "Pedido.Dias <= 999999"
            
            Listado.ReportFileName = "analisispedindicador.rpt"
            
            
    End Select
                        
    Listado.Connect = Connect()
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
    
    Desde.SetFocus
    PrgAnalisisPedido.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Auxiliar
End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "Resumido"
    Tipo.AddItem "Analitico C/Retrasos"
    Tipo.AddItem "Analitico Total"
    Tipo.AddItem "Analitico C/Retrasos por Concepto Produccion"
    Tipo.AddItem "Analitico de Entregas"
    Tipo.AddItem "Analitico de Entregas Analitico"
    Rem Tipo.AddItem "Analitico C/Retrasos por Concepto Expedicion"
    Tipo.AddItem "Indicador de incidencia de Falta de M.P en lo retrasos"
    
    Tipo.ListIndex = 0
    
    TipoFecha.Clear
    
    TipoFecha.AddItem "Ultima Fecha Programada"
    TipoFecha.AddItem "Fecha Original Pactada"
    
    TipoFecha.ListIndex = 0
    
    TipoFechaII.Clear
    
    TipoFechaII.AddItem "Fecha de Factura"
    TipoFechaII.AddItem "Fecha de Hoja de Ruta"
    
    TipoFechaII.ListIndex = 0

    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgAnalisisPedido.Caption = "Analisis de Cumplimiento de Pedidos de Venta :  " + !Nombre
        End If
    End With
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
End Sub

