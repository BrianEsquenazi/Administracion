VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgConsumoTerII 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de Analisis de Consumo de Productos Terminados"
   ClientHeight    =   6180
   ClientLeft      =   2085
   ClientTop       =   1500
   ClientWidth     =   8085
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   8085
   Begin VB.Frame Frame2 
      Height          =   3015
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   5295
      Begin VB.CommandButton Consulta 
         Caption         =   "Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3840
         TabIndex        =   16
         Top             =   450
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   1800
         TabIndex        =   11
         Top             =   600
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
         Left            =   1800
         TabIndex        =   0
         Top             =   240
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   2160
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
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   2160
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
         Left            =   3840
         TabIndex        =   8
         Top             =   960
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
         Left            =   3840
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin MSMask.MaskEdBox HastaFecha 
         Height          =   300
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox DesdeFecha 
         Height          =   300
         Left            =   1800
         TabIndex        =   13
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
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
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Prod.Term."
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Prod.Term."
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WConsumoTerII.rpt"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   840
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
      Height          =   2160
      ItemData        =   "consumoterii.frx":0000
      Left            =   120
      List            =   "consumoterii.frx":0007
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   6840
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgConsumoTerII"
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
Dim XParam As String
Dim Vector(10000, 7) As String
Private XLote(100, 7) As String
Private WCantidad As Double
Private WSaldo As Double

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    WDesde = Right$(DesdeFecha.Text, 4) + Mid$(DesdeFecha.Text, 4, 2) + Left$(DesdeFecha.Text, 2)
    WHasta = Right$(HastaFecha.Text, 4) + Mid$(HastaFecha.Text, 4, 2) + Left$(HastaFecha.Text, 2)
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    WVarios = "del " + DesdeFecha.Text + " al " + HastaFecha.Text

    Da = 0
    With rstFichaTer
        .Index = "Terminado"
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
    spTerminado = "ListaTerminadoDesdeHasta" + XParam
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
                WInicial = rstTerminado!Inicial
                WFechaCierre = IIf(IsNull(rstTerminado!FechaCierre), "00/00/0000", rstTerminado!FechaCierre)
                WOrdFechaCierre = IIf(IsNull(rstTerminado!OrdFechaCierre), "00000000", rstTerminado!OrdFechaCierre)
                WStock = rstTerminado!Inicial + rstTerminado!Entradas - rstTerminado!Salidas
                
                With rstFichaTer
                        .AddNew
                        !Terminado = WTerminado
                        !Fecha = WFechaCierre
                        !FechaOrd = "00000000"
                        !Tipo = 0
                        !Numero = 0
                        !Inicial = WStock
                        !Entrada = 0
                        !Salida = 0
                        !Observaciones = ""
                        !Lista1 = ""
                        !Lista2 = ""
                        !Titulo = WAuxiliar
                        .Update
                End With
                
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
    Renglon = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM Estadistica"
    Sql3 = " Where Estadistica.Articulo >= " + "'" + Desde.Text + "'"
    Sql4 = " and Estadistica.Articulo <= " + "'" + Hasta.Text + "'"
    Sql5 = " and Estadistica.OrdFecha >= " + "'" + WDesde + "'"
    Sql6 = " and Estadistica.OrdFecha <= " + "'" + WHasta + "'"
    Sql7 = " Order by Estadistica.OrdFecha"
    spEstadistica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7
    Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEstadistica.RecordCount > 0 Then
        With rstEstadistica
    
            .MoveFirst
            
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                        
                    XFec = Right$(rstEstadistica!Fecha, 4) + Mid$(rstEstadistica!Fecha, 4, 2) + Left$(rstEstadistica!Fecha, 2)
                    If XFec >= WDesde And XFec <= WHasta Then
                            
                        WTipo = rstEstadistica!Tipo
                        WTerminado = rstEstadistica!Articulo
                        WSalida = rstEstadistica!Cantidad
                        WFecha = rstEstadistica!Fecha
                        WNumero = rstEstadistica!Numero
                        WImpre1 = rstEstadistica!Cliente
                        WImpre2 = rstEstadistica!Cliente
                        
                        aa = rstEstadistica!Clave
                        
                        Erase XLote
                
                        XLote(1, 1) = IIf(IsNull(rstEstadistica!lote1), "", rstEstadistica!lote1)
                        XLote(1, 2) = IIf(IsNull(rstEstadistica!Canti1), "0", rstEstadistica!Canti1)
                        XLote(2, 1) = IIf(IsNull(rstEstadistica!lote2), "", rstEstadistica!lote2)
                        XLote(2, 2) = IIf(IsNull(rstEstadistica!Canti2), "0", rstEstadistica!Canti2)
                        XLote(3, 1) = IIf(IsNull(rstEstadistica!lote3), "", rstEstadistica!lote3)
                        XLote(3, 2) = IIf(IsNull(rstEstadistica!Canti3), "0", rstEstadistica!Canti3)
                        XLote(4, 1) = IIf(IsNull(rstEstadistica!lote4), "", rstEstadistica!lote4)
                        XLote(4, 2) = IIf(IsNull(rstEstadistica!Canti4), "0", rstEstadistica!Canti4)
                        XLote(5, 1) = IIf(IsNull(rstEstadistica!lote5), "", rstEstadistica!lote5)
                        XLote(5, 2) = IIf(IsNull(rstEstadistica!Canti5), "0", rstEstadistica!Canti5)
                    
                        If XLote(1, 2) = 0 Then
                            XLote(1, 2) = rstEstadistica!Cantidad
                        End If
                
                        For x = 1 To 5
                
                            If XLote(x, 2) <> 0 Then
                
                                WSalida = XLote(x, 2)
                                WLote = XLote(x, 1)
                                        
                                With rstFichaTer
                                    .AddNew
                                    !Terminado = WTerminado
                                    !Fecha = WFecha
                                    !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                    !Tipo = 0
                                    !Numero = WNumero
                                    !Inicial = 0
                                    If Val(WTipo) = 1 Then
                                        !Entrada = 0
                                        !Salida = WSalida
                                        !Lista1 = "Fact"
                                            Else
                                        !Entrada = 0
                                        !Salida = Abs(WSalida) * -1
                                        !Lista1 = "Devol"
                                    End If
                                    !Observaciones = ""
                                    !Lista2 = WImpre1 + " " + Left$(WImpre2, 23)
                                    !Lote = Val(WLote)
                                    !Saldo = 0
                                    !Titulo = WAuxiliar
                                    .Update
                                End With
                        
                            End If
                
                        Next x
                
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
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
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
                If XFec >= WDesde And XFec <= WHasta Then
                
                If rstHoja!Tipo = "T" Then
                
                    XLote(1, 1) = IIf(IsNull(rstHoja!lote1), "", rstHoja!lote1)
                    XLote(1, 2) = IIf(IsNull(rstHoja!Canti1), "", rstHoja!Canti1)
                    XLote(2, 1) = IIf(IsNull(rstHoja!lote2), "", rstHoja!lote2)
                    XLote(2, 2) = IIf(IsNull(rstHoja!Canti2), "", rstHoja!Canti2)
                    XLote(3, 1) = IIf(IsNull(rstHoja!lote3), "", rstHoja!lote3)
                    XLote(3, 2) = IIf(IsNull(rstHoja!Canti3), "", rstHoja!Canti3)
                        
                    If Val(XLote(1, 1)) = 0 Then
                        XLote(1, 1) = rstHoja!Lote
                        XLote(1, 2) = rstHoja!Cantidad
                    End If
                        
                    For Da = 1 To 3
                        
                        If Val(XLote(Da, 2)) <> 0 Then
                
                            WTerminado = rstHoja!Terminado
                            WCantidad = XLote(Da, 2)
                            WFecha = rstHoja!Fecha
                            WHoja = rstHoja!Hoja
                            WLote = XLote(Da, 1)
                
                            With rstFichaTer
                
                                .AddNew
                                !Terminado = WTerminado
                                !Fecha = WFecha
                                !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                                !Tipo = 0
                                !Numero = WHoja
                                !Inicial = 0
                                !Entrada = 0
                                !Salida = WCantidad
                                !Observaciones = ""
                                !Lista1 = "Hoja"
                                !Lista2 = ""
                                !Lote = WLote
                                !Saldo = 0
                                !Titulo = WAuxiliar
                                .Update
                            End With
                        End If
                    Next Da
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
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
                
                XFec = Right$(rstMovvar!Fecha, 4) + Mid$(rstMovvar!Fecha, 4, 2) + Left$(rstMovvar!Fecha, 2)
                If XFec >= WDesde And XFec <= WHasta Then
                
                If XFec >= WDesde And XFec <= WHasta And rstMovvar!Movi = "S" Then
                
                If rstMovvar!Tipo = "T" Then
                
                    WTerminado = rstMovvar!Terminado
                    WCantidad = rstMovvar!Cantidad
                    WFecha = rstMovvar!Fecha
                    WCodigo = rstMovvar!Codigo
                    WMovi = rstMovvar!Movi
                    WTipomov = Val(rstMovvar!Tipomov)
                    WObservaciones = rstMovvar!Observaciones
                    WLote = rstMovvar!Lote

                    With rstFichaTer
                
                        .AddNew
                        !Terminado = WTerminado
                        !Fecha = WFecha
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Tipo = 0
                        !Numero = WCodigo
                        !Inicial = 0
                        If WMovi = "E" Then
                            !Entrada = WCantidad
                            !Salida = 0
                                Else
                            !Entrada = 0
                            !Salida = WCantidad
                        End If
                        !Observaciones = ""
                        If WTipomov = 1 Or WTipomov = 2 Then
                            !Lista1 = "Mov.Var"
                                Else
                            !Lista1 = "Guia In"
                        End If
                        !Lista2 = Left$(WObservaciones, 30)
                        !Lote = WLote
                        !Saldo = 0
                        !Titulo = WAuxiliar
                        .Update
                    End With
                    
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
    
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
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
                
                XFec = Right$(rstMovguia!Fecha, 4) + Mid$(rstMovguia!Fecha, 4, 2) + Left$(rstMovguia!Fecha, 2)
                If XFec >= WDesde And XFec <= WHasta And rstMovguia!Movi = "S" Then
                
                If rstMovguia!Tipo = "T" Then
                
                    WTerminado = rstMovguia!Terminado
                    WCantidad = rstMovguia!Cantidad
                    WFecha = rstMovguia!Fecha
                    WCodigo = rstMovguia!Codigo
                    WMovi = rstMovguia!Movi
                    Rem WObservaciones = rstMovvar!Observaciones
                    WDestino = rstMovguia!Destino
                    WTipomov = rstMovguia!Tipomov
                    
                    If WMovi = "S" Then
                            Select Case WDestino
                                Case 1
                                    WObservaciones = "Envio a Surfactan"
                                Case 2
                                    WObservaciones = "Envio a Pellital"
                                Case 3
                                    WObservaciones = "Envio a Surfactan II"
                                Case 4
                                    WObservaciones = "Envio a Pellital II"
                                Case 5
                                    WObservaciones = "Envio a Surfactan III"
                                Case 6
                                    WObservaciones = "Envio a Surfactan IV"
                                Case 7
                                    WObservaciones = "Envio a Surfactan V"
                                Case 8
                                    WObservaciones = "Envio a Pellital V"
                                Case 9
                                    WObservaciones = "Envio a Pellital IV"
                                Case 10
                                    WObservaciones = "Envio a Surfactan VI"
                                Case 11
                                    WObservaciones = "Envio a Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Partida
                            WSaldo = 0
                            
                                Else
                                
                            Select Case WTipomov
                                Case 1
                                    WObservaciones = "Recepcion de Surfactan"
                                Case 2
                                    WObservaciones = "Recepcion de Pellital"
                                Case 3
                                    WObservaciones = "Recepcion de Surfactan II"
                                Case 4
                                    WObservaciones = "Recepcion de Pellital II"
                                Case 5
                                    WObservaciones = "Recepcion de Surfactan III"
                                Case 6
                                    WObservaciones = "Recepcion de Surfactan IV"
                                Case 7
                                    WObservaciones = "Recepcion de Surfactan V"
                                Case 8
                                    WObservaciones = "Recepcion de Pellital V"
                                Case 9
                                    WObservaciones = "Recepcion de Pellital IV"
                                Case 10
                                    WObservaciones = "Recepcion de Surfactan VI"
                                Case 11
                                    WObservaciones = "Recepcion de Surfactan VII"
                                Case Else
                            End Select
                            WLote = rstMovguia!Lote
                            WSaldo = rstMovguia!Saldo
                            Call Redondeo(WSaldo)
                            
                    End If
                        
                    With rstFichaTer
                
                        .AddNew
                        !Terminado = WTerminado
                        !Fecha = WFecha
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Tipo = 0
                        !Numero = WCodigo
                        !Inicial = 0
                        If WMovi = "E" Then
                            !Entrada = WCantidad
                            !Salida = 0
                                Else
                            !Entrada = 0
                            !Salida = WCantidad
                        End If
                        !Observaciones = ""
                        If !Numero > 900000 Then
                            !Lista1 = "Prestamo"
                            !Numero = !Numero - 900000
                                Else
                            !Lista1 = "Guia In"
                        End If
                        !Lista2 = Left$(WObservaciones, 30)
                        !Lote = WLote
                        !Saldo = WSaldo
                        !Titulo = WAuxiliar
                        .Update
                    End With
                    
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
    
    
    
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
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
                
                XFec = Right$(rstConsig!Fecha, 4) + Mid$(rstConsig!Fecha, 4, 2) + Left$(rstConsig!Fecha, 2)
                If XFec >= WDesde And XFec <= WHasta Then
                
                    WTerminado = rstConsig!Terminado
                    WCantidad = rstConsig!Cantidad - rstConsig!Facturado
                    WFecha = rstConsig!Fecha
                    WCodigo = rstConsig!Numero
                    WCliente = rstConsig!Cliente
                    WObservaciones = rstConsig!Observaciones
                    WLote = rstConsig!Lote
                    
                    If WCantidad <> 0 Then

                        With rstFichaTer
                            .AddNew
                            !Terminado = WTerminado
                            !Fecha = WFecha
                            !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                            !Tipo = 0
                            !Numero = WCodigo
                            !Inicial = 0
                            !Entrada = 0
                            !Salida = WCantidad
                            !Observaciones = WCliente
                            !Lista1 = "Rem.Con."
                            !Lista2 = Left$(WObservaciones, 30)
                            !Lote = WLote
                            !Saldo = 0
                            !Titulo = WAuxiliar
                            .Update
                        End With
                        
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
    
    XParam = "'" + Desde.Text + "','" _
                 + Hasta.Text + "'"
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
                
                XFec = Right$(rstMovlab!Fecha, 4) + Mid$(rstMovlab!Fecha, 4, 2) + Left$(rstMovlab!Fecha, 2)
                If XFec >= WDesde And XFec <= WHasta And rstMovlab!Movi = "S" Then
                
                If rstMovlab!Tipo = "T" Then
                
                    WTerminado = rstMovlab!Terminado
                    WCantidad = rstMovlab!Cantidad
                    WFecha = rstMovlab!Fecha
                    WCodigo = rstMovlab!Codigo
                    WMovi = rstMovlab!Movi
                    WTipomov = rstMovlab!Tipomov
                    WObservaciones = rstMovlab!Observaciones
                    WLote = rstMovlab!Lote

                    With rstFichaTer
                
                        .AddNew
                        !Terminado = WTerminado
                        !Fecha = WFecha
                        !FechaOrd = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
                        !Tipo = 0
                        !Numero = WCodigo
                        !Inicial = 0
                        If WMovi = "E" Then
                            !Entrada = WCantidad
                            !Salida = 0
                                Else
                            !Entrada = 0
                            !Salida = WCantidad
                        End If
                        !Observaciones = ""
                        !Lista1 = "Mov.Lab"
                        !Lista2 = Left$(WObservaciones, 30)
                        !Lote = WLote
                        !Saldo = 0
                        !Titulo = WAuxiliar
                        .Update
                    End With
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
    
    
    
    Da = 0
    With rstFichaTer
        .Index = "Terminado"
        .Seek ">=", ""
        If .NoMatch = False Then
            Do
                .Edit
                
                WTerminado = !Terminado
                WDescripcion = ""
                spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    WDescripcion = rstTerminado!Descripcion
                    rstTerminado.Close
                End If
                !Descripcion = WDescripcion
                !Observaciones = WVarios
                
                .Update
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    Listado.WindowTitle = "Listado de Consumo de Producto Terminado "
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Rem Listado.GroupSelectionFormula = "{FichaTer.Terminado} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
   
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If

    Listado.DataFiles(0) = WEmpresa + "auxi.mdb"
    
    Listado.Action = 1
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    With rstFichaTer
        .Close
    End With
    DbsEmpresa.Close
    
    Desde.SetFocus
    PrgConsumoTerII.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.Text = Desde.Text
        Hasta.SetFocus
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

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        DesdeFecha.SetFocus
    End If
End Sub

Private Sub DesdeFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaFecha.SetFocus
    End If
End Sub

Private Sub HastaFecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgConsumoTerII.Caption = "Listado de Analisis de Consumo de Productos Terminados :  " + !Nombre
        End If
    End With
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    DesdeFecha.Text = "  /  /    "
    HastaFecha.Text = "  /  /    "
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


