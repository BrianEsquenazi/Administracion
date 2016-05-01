VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgEtiVerde 
   Caption         =   "Impresion de Etiquetas Verdes"
   ClientHeight    =   5685
   ClientLeft      =   1170
   ClientTop       =   1485
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   10425
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   9855
      Begin VB.CommandButton Command1 
         Caption         =   "Etiqueta SGS"
         Height          =   735
         Left            =   8280
         TabIndex        =   25
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Informe 
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
         Left            =   5760
         MaxLength       =   20
         TabIndex        =   22
         Text            =   "  "
         Top             =   1680
         Width           =   2055
      End
      Begin VB.ComboBox TipoEtiqueta 
         Height          =   315
         Left            =   4560
         TabIndex        =   21
         Top             =   3240
         Width           =   3135
      End
      Begin VB.TextBox Partida 
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
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "  "
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Fecha 
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
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   18
         Text            =   " "
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox Codigo 
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
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   17
         Text            =   " "
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox Tipo 
         BackColor       =   &H0080FFFF&
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
         Left            =   5280
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Baja 
         Caption         =   "  Limpia Etiquetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   14
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox Cantidad 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   12
         Text            =   " "
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox Kilos 
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   11
         Text            =   " "
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Lote 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
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
         Left            =   2280
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "  "
         Top             =   480
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
         Height          =   495
         Left            =   8520
         TabIndex        =   6
         Top             =   1560
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
         Height          =   495
         Left            =   8520
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   120
         X2              =   7800
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   7800
         X2              =   7800
         Y1              =   960
         Y2              =   360
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   120
         Y1              =   960
         Y2              =   360
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   7800
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   4080
         TabIndex        =   24
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Informe"
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
         Left            =   4800
         TabIndex        =   23
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Partida"
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
         Left            =   240
         TabIndex        =   20
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label DesCodigo 
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
         Height          =   255
         Left            =   3960
         TabIndex        =   13
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de Etiquetas"
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
         Left            =   240
         TabIndex        =   10
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Kilos Envase"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Lote"
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
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "M.P./P.T."
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
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "eti1.rpt"
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
      Left            =   5760
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   300
      Left            =   4680
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   4680
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgEtiVerde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WLote As String
Private WCantidad As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim XParam As String
Dim Da As Integer
Dim XMes As String
Dim XAno As String
Dim Empe(12, 10) As String
Private WImpreadi As String
Private WClase As String
Private WIntervencion As String
Private WNaciones As String
Private WEmbalaje As String
Dim ZVencimiento As String

Dim ZZPartiOri As String

Dim ZImpre(1000) As String
Dim ZImpreI(1000) As String
Dim ZImpreII(1000) As String
Dim ZImpreIII(1000) As String

Dim ZLugarImpre As Integer
Dim ZLugarImpreI As Integer
Dim ZLugarImpreII As Integer
Dim ZLugarImpreIII As Integer
Dim ZZLogo(100) As Integer
Dim ZZImpreFrase(100) As String

Private Sub Acepta_Click()

    On Error GoTo WError
    
    Salida = "N"
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Imprsion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
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
    
        WNeto = Val(Kilos.Text)
        
        ZMarca = 0
        ZCantidad = Int(Val(Cantidad.Text) / 2)
        If ZCantidad * 2 <> Val(Cantidad.Text) Then
            ZCantidad = ZCantidad + 1
            ZMarca = 1
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            rstArticulo.Close
        End If
        
        Codigo.Text = UCase(Codigo.Text)
        
        If Tipo.ListIndex = 1 Then
              XCodigo = Val(Mid$(Codigo.Text, 4, 5))
                  If XCodigo >= 0 And XCodigo <= 999 Then
                     TipoEtiqueta.ListIndex = 2
                  End If
                  If XCodigo >= 11000 And XCodigo <= 12999 Then
                      TipoEtiqueta.ListIndex = 2
                  End If
             Else
                    Rem BY NAN
                         XCodigo = Val(Mid$(Codigo.Text, 4, 5))
                                         
                                    If Left$(Codigo.Text, 2) = "DY" Or Left$(Codigo.Text, 2) = "DS" Then
                                       TipoEtiqueta.ListIndex = 2
                                     End If
                                       If Left$(Codigo.Text, 2) = "CO" Then
                                              If XCodigo >= 0 And XCodigo <= 999 Then
                                                  TipoEtiqueta.ListIndex = 2
                                                  End If
                                                  If XCodigo >= 11000 And XCodigo <= 12999 Then
                                                    TipoEtiqueta.ListIndex = 2
                                                    End If
                                       End If
                                 
            
          End If
        
        Select Case TipoEtiqueta.ListIndex
            Case 0
                If Tipo.ListIndex = 0 Then
        
                    If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
        
                        ZMeses = 0
                        WTipoeti = ""
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                            ZMeses = rstArticulo!Meses
                            WTipoeti = IIf(IsNull(rstArticulo!TipoEti), "", rstArticulo!TipoEti)
                            rstArticulo.Close
                        End If
        
                        WMes = Val(Mid$(Fecha.Text, 4, 2))
                        WAno = Val(Right$(Fecha.Text, 4))
                        For ZCiclo = 1 To ZMeses
                            WMes = WMes + 1
                            If WMes > 12 Then
                                WAno = WAno + 1
                                WMes = 1
                            End If
                        Next ZCiclo
            
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        If Val(Left$(Fecha.Text, 2)) <= 30 Then
                            If Val(XMes) = 2 And Val(Left$(Fecha.Text, 2)) > 28 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = Left$(Fecha.Text, 3) + XMes + "/" + XAno
                            End If
                                Else
                            If Val(XMes) = 2 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = "30/" + XMes + "/" + XAno
                            End If
                        End If
            
                    End If
            
                End If
        
        
                With rstEtiqueta
                    For Da = 1 To ZCantidad
                        .Index = "Codigo"
                        .AddNew
                
                        WLote = Lote.Text
                        Call Ceros(WLote, 6)
                
                        WCantidad = Kilos.Text
                        Call Ceros(WCantidad, 4)
                
                        ZDa = Int((Da - 1) / 2)
                
                        !Codigo = Da
                        !Terminado = Codigo.Text
                        !Lote = WLote
                        !Cliente = ""
                        !Cantidad = Val(Kilos.Text)
                        !Nombre = "Fec.Lau.: " + Fecha.Text
                        If ZVencimiento <> "00/00/0000" Then
                            !Impre1 = "Fec.Rea.:" + ZVencimiento
                                Else
                            !Impre1 = ""
                        End If
                        !Conservacion = !Impre1
                        
                        Rem XTipoPro = ""
                        Rem If Left$(WTerminado, 2) = "PT" Then
                        Rem     If XCodigo >= 0 And XCodigo <= 999 Then
                        Rem         XTipoPro = "CO"
                        Rem             Else
                        Rem         If XCodigo >= 11000 And XCodigo <= 12999 Then
                        Rem             XTipoPro = "CO"
                        Rem         End If
                        Rem     End If
                        Rem End If
                        Rem If XTipoPro = "CO" Then
                        Rem     !Nombre = ""
                        Rem     !Impre1 = ""
                        Rem     !Conservacion = !Impre1
                        Rem End If
                        
                        
                        
                        !Razon = "L:" + Partida.Text
                        !DirEntrega = Kilos.Text + " Kgs."
                        !Clase = WClase
                        !Intervencion = WIntervencion
                        !Naciones = WNaciones
                        !Embalaje = WEmbalaje
                        !Bruto = 0
                        If Da = ZCantidad And ZMarca = 1 Then
                            !Bruto = 1
                        End If
                        !Neto = ZDa
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            !Observaciones = "CONTROL CALIDAD"
                                Else
                            !Observaciones = "C.C.   PELLITAL"
                        End If
                        
                        !Elaboracion = WTipoeti
                        
                        .Update
                    Next Da
                End With

                Listado.WindowTitle = "Emision de Etiquetas"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
                
                Select Case Mid$(WClase, 1, 1)
                    Case "3"
                        Listado.ReportFileName = "WEtiVerde3.rpt"
                    Case "4"
                        Listado.ReportFileName = "WEtiVerde4.rpt"
                    Case "5"
                        Listado.ReportFileName = "WEtiVerde5.rpt"
                    Case "6"
                        Listado.ReportFileName = "WEtiVerde6.rpt"
                    Case "8"
                        Listado.ReportFileName = "WEtiVerde8.rpt"
                    Case "9"
                        Listado.ReportFileName = "WEtiVerde9.rpt"
                    Case Else
                        Listado.ReportFileName = "WEtiVerde.rpt"
                End Select

                Rem Listado.ReportFileName = "WEtiVerde.rpt"
                Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
                Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
                Rem Listado.Connect = Connect()
    
                Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
                Listado.Destination = 1
                Rem Listado.Destination = 0
                Listado.PrinterCopies = 1
                Listado.Action = 1
                
            Case 1
                If Tipo.ListIndex = 0 Then
        
                    If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
        
                        ZMeses = 0
                        WTipoeti = ""
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                            ZMeses = rstArticulo!Meses
                            WTipoeti = IIf(IsNull(rstArticulo!TipoEti), "", rstArticulo!TipoEti)
                            rstArticulo.Close
                        End If
        
                        WMes = Val(Mid$(Fecha.Text, 4, 2))
                        WAno = Val(Right$(Fecha.Text, 4))
                        For ZCiclo = 1 To ZMeses
                            WMes = WMes + 1
                            If WMes > 12 Then
                                WAno = WAno + 1
                                WMes = 1
                            End If
                        Next ZCiclo
            
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        If Val(Left$(Fecha.Text, 2)) <= 30 Then
                            If Val(XMes) = 2 And Val(Left$(Fecha.Text, 2)) > 28 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = Left$(Fecha.Text, 3) + XMes + "/" + XAno
                            End If
                                Else
                            If Val(XMes) = 2 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = "30/" + XMes + "/" + XAno
                            End If
                        End If
            
                    End If
            
                End If
        
        
                With rstEtiqueta
                    For Da = 1 To ZCantidad
                        .Index = "Codigo"
                        .AddNew
                
                        WLote = Lote.Text
                        Call Ceros(WLote, 6)
                
                        WCantidad = Kilos.Text
                        Call Ceros(WCantidad, 4)
                
                        ZDa = Int((Da - 1) / 2)
                
                        !Codigo = Da
                        !Terminado = Codigo.Text
                        !Lote = WLote
                        !Cliente = ""
                        !Cantidad = Val(Kilos.Text)
                        !Nombre = Left$(DesCodigo.Caption, 35)
                        If ZVencimiento <> "00/00/0000" Then
                            !DirEntrega = "Fecha Reanalisis : " + ZVencimiento
                                Else
                            !DirEntrega = ""
                        End If
                        !Razon = "Lote : " + Partida.Text
                        Rem !DirEntrega = "Cantidad por Bulto : " + Kilos.Text + " Kgs."
                        Rem !DirEntrega = ""
                        !Conservacion = "Fecha de Ingreso : " + Fecha.Text
                        !Clase = WClase
                        !Intervencion = WIntervencion
                        !Naciones = WNaciones
                        !Embalaje = WEmbalaje
                        !Bruto = 0
                        If Da = ZCantidad And ZMarca = 1 Then
                            !Bruto = 1
                        End If
                        !Neto = ZDa
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            !Observaciones = "CONTROL CALIDAD"
                                Else
                            !Observaciones = "C.C.   PELLITAL"
                        End If
                        !ConservacionII = DesCodigo.Caption
                        !Elaboracion = WTipoeti
                        .Update
                    Next Da
                End With

                Listado.WindowTitle = "Emision de Etiquetas"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
        
                Select Case Mid$(WClase, 1, 1)
                    Case "3"
                        Listado.ReportFileName = "WEtiVerdeFarma3.rpt"
                    Case "5"
                        Listado.ReportFileName = "WEtiVerdeFarma5.rpt"
                    Case "6"
                        Listado.ReportFileName = "WEtiVerdeFarma6.rpt"
                    Case "8"
                        Listado.ReportFileName = "WEtiVerdeFarma8.rpt"
                    Case "9"
                        Listado.ReportFileName = "WEtiVerdeFarma9.rpt"
                    Case Else
                      Listado.ReportFileName = "WEtiVerdeFarma.rpt"
                   Rem Listado.ReportFileName = "WEtinan.rpt"
                
                End Select

                Rem Listado.ReportFileName = "WEtiVerde.rpt"
                Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
                Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
                Rem Listado.Connect = Connect()
    
                Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
                Listado.Destination = 1
                Rem Listado.Destination = 0
                Listado.PrinterCopies = 1
                Listado.Action = 1
                
            Case 2
                
                
                If Tipo.ListIndex = 0 Then
        
                    If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
        
                        ZMeses = 0
                        WTipoeti = ""
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Articulo"
                        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                        spArticulo = ZSql
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                            ZMeses = rstArticulo!Meses
                            WTipoeti = IIf(IsNull(rstArticulo!TipoEti), "", rstArticulo!TipoEti)
                            rstArticulo.Close
                        End If
        
                        WMes = Val(Mid$(Fecha.Text, 4, 2))
                        WAno = Val(Right$(Fecha.Text, 4))
                        For ZCiclo = 1 To ZMeses
                            WMes = WMes + 1
                            If WMes > 12 Then
                                WAno = WAno + 1
                                WMes = 1
                            End If
                        Next ZCiclo
            
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        If Val(Left$(Fecha.Text, 2)) <= 30 Then
                            If Val(XMes) = 2 And Val(Left$(Fecha.Text, 2)) > 28 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = Left$(Fecha.Text, 3) + XMes + "/" + XAno
                            End If
                                Else
                            If Val(XMes) = 2 Then
                                ZVencimiento = "28/" + XMes + "/" + XAno
                                    Else
                                ZVencimiento = "30/" + XMes + "/" + XAno
                            End If
                        End If
            
                    End If
            
                End If
        
                Rem BY NAN 19-6-2015 NO DEBE SALIR FECHA REAL NI LAUDO IMPRESO
                ZVencimiento = "00/00/0000"
                Rem BY NAN
        
                With rstEtiqueta
                    For Da = 1 To ZCantidad
                        .Index = "Codigo"
                        .AddNew
                
                        WLote = Lote.Text
                        Call Ceros(WLote, 6)
                
                        WCantidad = Kilos.Text
                        Call Ceros(WCantidad, 4)
                
                        ZDa = Int((Da - 1) / 2)
                
                        !Codigo = Da
                        If Tipo.ListIndex = 0 Then
                            !Terminado = Codigo.Text
                                Else
                            !Terminado = Mid$(Codigo.Text, 4, 20)
                        End If
                        !Lote = WLote
                        !Cliente = ""
                        !Cantidad = Val(Kilos.Text)
                        !Nombre = "Fec.Lau.: " + Fecha.Text
                        If ZVencimiento <> "00/00/0000" Then
                            !Impre1 = "Fec.Rea.:" + ZVencimiento
                                Else
                            !Impre1 = ""
                        End If
                         
                        !Nombre = ""
                        !Impre1 = WTipoeti
                        
                        
                        !Conservacion = !Impre1
                        
                    
                        
                        Sql1 = "Select *"
                        Sql2 = " FROM PrueArt"
                        Sql3 = " WHERE Lote = " + "'" + Partida.Text + "'"
                        spPrueart = Sql1 + Sql2 + Sql3
                        Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
                        If rstPrueart.RecordCount > 0 Then
                            !Razon = Partida.Text
                            rstPrueart.Close
                                Else
                            !Razon = ZZPartiOri
                            !Nombre = Partida.Text
                        End If
                        
                        !DirEntrega = Kilos.Text + " Kgs."
                        !Clase = WClase
                        !Intervencion = WIntervencion
                        !Naciones = WNaciones
                        !Embalaje = WEmbalaje
                        !Bruto = 0
                        If Da = ZCantidad And ZMarca = 1 Then
                            !Bruto = 1
                        End If
                        !Neto = ZDa
                        If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                            !Observaciones = "CONTROL INTERNO"
                                Else
                            !Observaciones = "C.C.   PELLITAL"
                        End If
                        !Elaboracion = WTipoeti
                        .Update
                    Next Da
                End With

                Listado.WindowTitle = "Emision de Etiquetas"
                Listado.WindowTop = 0
                Listado.WindowLeft = 0
                Listado.WindowWidth = Screen.Width
                Listado.WindowHeight = Screen.Height
                
                Select Case Mid$(WClase, 1, 1)
                    Case "3"
                        Listado.ReportFileName = "WEtiVerde3.rpt"
                    Case "4"
                        Listado.ReportFileName = "WEtiVerde4.rpt"
                    Case "5"
                        Listado.ReportFileName = "WEtiVerde5.rpt"
                    Case "6"
                        Listado.ReportFileName = "WEtiVerde6.rpt"
                    Case "8"
                        Listado.ReportFileName = "WEtiVerde8.rpt"
                    Case "9"
                        Listado.ReportFileName = "WEtiVerde9.rpt"
                    Case Else
                        Listado.ReportFileName = "WEtiVerde.rpt"
                Rem Listado.ReportFileName = "WEtinan.rpt"
                
                End Select

                Rem Listado.ReportFileName = "WEtiVerde.rpt"
                Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
                Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
                Rem Listado.Connect = Connect()
    
                Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
    
                Listado.Destination = 1
                Rem Listado.Destination = 0
                Listado.PrinterCopies = 1
                Listado.Action = 1
            
            Case Else
                       
        
                Rem ******************by nan 03-12-2014 eti ni solo codigo y Kilos
        
                WTipoeti = ""
                spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DesCodigo.Caption = rstArticulo!Descripcion
                    WImpreadi = ""
                    WClase = ""
                    WIntervencion = ""
                    WNaciones = ""
                    WEmbalaje = ""
                    WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                    WIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                    WNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
                    WEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
                    WTipoeti = IIf(IsNull(rstArticulo!TipoEti), "", rstArticulo!TipoEti)
                    rstArticulo.Close
                End If
                                       
                      
                If Len(DesCodigo.Caption) > 30 Then
                    For Da = 31 To 1 Step -1
                        If Mid$(DesCodigo.Caption, Da, 1) = Space$(1) Then
                            ZZNombre = Mid$(DesCodigo.Caption, 1, Da)
                            ZZNombreII = Mid$(DesCodigo.Caption, Da + 1, 100)
                            Exit For
                        End If
                    Next Da
                        Else
                    ZZNombre = DesCodigo.Caption
                    ZZNombreII = ""
                End If
                      
                      
                      
                      
                      
                      
                      
                      
                      
                      
                       With rstEtiqueta
                          For Da = 1 To ZCantidad
                                .Index = "Codigo"
                                .AddNew
                
                          Rem  WLote = Lote.Text
                              Rem  Call Ceros(WLote, 6)
                    
                              WCantidad = Kilos.Text
                            Call Ceros(WCantidad, 4)
                    
                            ZDa = Int((Da - 1) / 2)
                    
                            !Codigo = Da
                                 !Terminado = Codigo.Text
                            Rem !Lote = WLote
                             !Cliente = ""
                             !Cantidad = Val(Kilos.Text)
                             !Nombre = Left$(ZZNombre, 30)
                            
                             Rem !Nombre = "Fec.Lau.: " + Fecha.Text
                            Rem  If ZVencimiento <> "00/00/0000" Then
                               Rem   !Impre1 = "Fec.Rea.:" + ZVencimiento
                               Rem       Else
                               Rem   !Impre1 = ""
                            Rem  End If
                              !Conservacion = IIf(IsNull(!Impre1), "", !Impre1)
                            Rem  !Razon = "L:" + Partida.Text
                              !DirEntrega = Kilos.Text + " Kgs."
                              !Clase = WClase
                              !Intervencion = WIntervencion
                              !Naciones = WNaciones
                              !Embalaje = WEmbalaje
                              !Bruto = 0
                              If Da = ZCantidad And ZMarca = 1 Then
                                  !Bruto = 1
                              End If
                              !Neto = ZDa
                                !Elaboracion = WTipoeti
                                Rem  If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                                 Rem    !Observaciones = "CONTROL CALIDAD"
                                 Rem        Else
                                 Rem    !Observaciones = "C.C.   PELLITAL"
                                 Rem  End If
                        .Update
                    Next Da
                End With

                        Listado.WindowTitle = "Emision de Etiquetas"
                        Listado.WindowTop = 0
                        Listado.WindowLeft = 0
                        Listado.WindowWidth = Screen.Width
                        Listado.WindowHeight = Screen.Height
                        
                        Select Case Mid$(WClase, 1, 1)
                            Case "3"
                                Listado.ReportFileName = "WEtiVerde3.rpt"
                            Case "4"
                                Listado.ReportFileName = "WEtiVerde4.rpt"
                            Case "5"
                                Listado.ReportFileName = "WEtiVerde5.rpt"
                            Case "6"
                                Listado.ReportFileName = "WEtiVerde6.rpt"
                            Case "8"
                                Listado.ReportFileName = "WEtiVerde8.rpt"
                            Case "9"
                                Listado.ReportFileName = "WEtiVerde9.rpt"
                            Case Else
                                Listado.ReportFileName = "WEtiVerde.rpt"
                        End Select
        
                        Rem Listado.ReportFileName = "WEtiVerde.rpt"
                        Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
                        Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
                        Rem Listado.Connect = Connect()
            
                        Listado.DataFiles(0) = Wempresa + "Auxi.mdb"
            
                        Listado.Destination = 1
                        Listado.PrinterCopies = Cantidad
                        Listado.Action = 1
        
        
        
        
                  Rem************* fin 03-12-2014   by nan fin
          
        
        End Select
    
        Da = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", Da
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
    
    End If
    
  Lote.Text = ""
    Codigo.Text = ""
    DesCodigo.Caption = ""
    Fecha.Text = ""
    Kilos.Text = ""
    Cantidad.Text = ""
    TipoEtiqueta.ListIndex = 0
    
    Partida.Visible = True
    Partida.Text = ""
             
             Informe.Visible = True
              Informe.Text = ""
              Lote.Enabled = True
              Lote.Visible = True
              Codigo.Locked = True
              Fecha.Visible = True
              Lote.SetFocus
    
    
    
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_Click()

    With rstEmpresa
        .Close
    End With
    PrgEtiVerde.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Baja_Click()
    Da = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", Da
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
    
End Sub

Private Sub Codigo_Change()
If TipoEtiqueta.ListIndex = 3 Then
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DesCodigo.Caption = rstArticulo!Descripcion
                Else
                

                End If



End If
 
End Sub



Private Sub Codigo_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then

 spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DesCodigo.Caption = rstArticulo!Descripcion
                Kilos.SetFocus
                
                Else
                Codigo.Text = ""
                
                DesCodigo.Caption = ""
       End If
       





End If
End Sub

Private Sub Command1_Click()

    Rem On Error GoTo WError
    
    Erase ZImpre
    Erase ZImpreI
    Erase ZImpreII
    Erase ZImpreIII
    
    ZLugarImpre = 0
    ZLugarImpreI = 0
    ZLugarImpreII = 0
    ZLugarImpreIII = 0
    
    Da = 0
    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
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
    
    Da = 0
    With rstEtiquetaIII
        .Index = "Codigo"
        .Seek ">=", Da
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
    
    Da = 0
    With rstEtiquetaIV
        .Index = "Codigo"
        .Seek ">=", Da
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
    
    
    Salida = "N"
    
    Da = 0
    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    Da = 0
    With rstEtiquetaIII
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    Da = 0
    With rstEtiquetaIV
        .Index = "Codigo"
        .Seek ">=", Da
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Impresion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
    
        Da = 0
        With rstEtiquetaII
            .Index = "Codigo"
            .Seek ">=", Da
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
    
        WNeto = Val(Kilos.Text)
        
        ZCantidad = Val(Cantidad.Text)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Articulo"
        ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            rstArticulo.Close
        End If
        
        Codigo.Text = UCase(Codigo.Text)
        
        If Tipo.ListIndex = 1 Then
            XCodigo = Val(Mid$(Codigo.Text, 4, 5))
            If XCodigo >= 0 And XCodigo <= 999 Then
                TipoEtiqueta.ListIndex = 2
            End If
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                TipoEtiqueta.ListIndex = 2
            End If
        Else
        
        Rem BY NAN
        XCodigo = Val(Mid$(Codigo.Text, 4, 5))
                                         
        If Left$(Codigo.Text, 2) = "DY" Or Left$(Codigo.Text, 2) = "DS" Then
            TipoEtiqueta.ListIndex = 2
        End If
        If Left$(Codigo.Text, 2) = "CO" Then
            If XCodigo >= 0 And XCodigo <= 999 Then
                TipoEtiqueta.ListIndex = 2
            End If
            If XCodigo >= 11000 And XCodigo <= 12999 Then
                TipoEtiqueta.ListIndex = 2
            End If
        End If
                                 
            
    End If
        
    If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
        
            ZMeses = 0

            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Articulo"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spArticulo = ZSql
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                ZMeses = rstArticulo!Meses
                rstArticulo.Close
            End If

            WMes = Val(Mid$(Fecha.Text, 4, 2))
            WAno = Val(Right$(Fecha.Text, 4))
            For ZCiclo = 1 To ZMeses
                WMes = WMes + 1
                If WMes > 12 Then
                    WAno = WAno + 1
                    WMes = 1
                End If
            Next ZCiclo

            XMes = Str$(WMes)
            XAno = Str$(WAno)
            Call Ceros(XMes, 2)
            Call Ceros(XAno, 4)
            If Val(Left$(Fecha.Text, 2)) <= 30 Then
                If Val(XMes) = 2 And Val(Left$(Fecha.Text, 2)) > 28 Then
                    ZVencimiento = "28/" + XMes + "/" + XAno
                        Else
                    ZVencimiento = Left$(Fecha.Text, 3) + XMes + "/" + XAno
                End If
                    Else
                If Val(XMes) = 2 Then
                    ZVencimiento = "28/" + XMes + "/" + XAno
                        Else
                    ZVencimiento = "30/" + XMes + "/" + XAno
                End If
            End If

        End If

    End If
    
    
    
    
    
    
            
            
            
            
        XEmpresa = Wempresa
        Select Case Val(XEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    Wempresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2, 4, 8, 9
                    Wempresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
        End Select
                
                
        For Ciclo = 1 To 999
        
            Auxi = Ciclo
            Call Ceros(Auxi, 3)
            
            ZZClave = Codigo.Text + Auxi
        
            Sql1 = "Select *"
            Sql2 = " FROM DatosEtiquetaMp"
            Sql3 = " Where DatosEtiquetaMp.Clave = " + "'" + ZZClave + "'"
            spDatosEtiquetaMp = Sql1 + Sql2 + Sql3
            Set rstDatosEtiquetaMp = db.OpenRecordset(spDatosEtiquetaMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstDatosEtiquetaMp.RecordCount > 0 Then
        
                ZZPalabra = rstDatosEtiquetaMp!Palabra
                ZZLogo(1) = rstDatosEtiquetaMp!pictograma1
                ZZLogo(2) = rstDatosEtiquetaMp!pictograma2
                ZZLogo(3) = rstDatosEtiquetaMp!pictograma3
                ZZLogo(4) = rstDatosEtiquetaMp!pictograma4
                ZZLogo(5) = rstDatosEtiquetaMp!pictograma5
                ZZLogo(6) = rstDatosEtiquetaMp!pictograma6
                ZZLogo(7) = rstDatosEtiquetaMp!pictograma7
                ZZLogo(8) = rstDatosEtiquetaMp!pictograma8
                ZZLogo(9) = rstDatosEtiquetaMp!pictograma9
                
                Select Case rstDatosEtiquetaMp!Tipo
                    Case 1
                        If Trim(rstDatosEtiquetaMp!descripcion1h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiquetaMp!descripcion1h)
                        End If
                        If Trim(rstDatosEtiquetaMp!descripcion2h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiquetaMp!descripcion2h)
                        End If
                        If Trim(rstDatosEtiquetaMp!descripcion3h) <> "" Then
                            ZLugarImpreI = ZLugarImpreI + 1
                            ZImpreI(ZLugarImpreI) = Trim(rstDatosEtiquetaMp!descripcion3h)
                        End If
                        
                    Case 2
                        If Trim(rstDatosEtiquetaMp!descripcion1p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiquetaMp!descripcion1p)
                        End If
                        If Trim(rstDatosEtiquetaMp!descripcion2p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiquetaMp!descripcion2p)
                        End If
                        If Trim(rstDatosEtiquetaMp!descripcion3p) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiquetaMp!descripcion3p)
                        End If
                        If Trim(rstDatosEtiquetaMp!Observaciones) <> "" Then
                            ZLugarImpreII = ZLugarImpreII + 1
                            ZImpreII(ZLugarImpreII) = Trim(rstDatosEtiquetaMp!Observaciones)
                        End If
                        
                    Case 3
                        If Trim(rstDatosEtiquetaMp!denominacion) <> "" Then
                            ZLugarImpreIII = ZLugarImpreIII + 1
                            ZImpreIII(ZLugarImpreIII) = Trim(rstDatosEtiquetaMp!denominacion)
                        End If
                        
                    Case Else
                    
                End Select
                
                
                    Else
                    
                Exit For
                
            End If
                    
        Next Ciclo
        
    
        Rem For Ciclo = 1 To 99
        Rem     If Trim(ZImpreI(Ciclo)) <> "" Then
        Rem         Sql1 = "Select *"
        Rem         Sql2 = " FROM FraseH"
        Rem         Sql3 = " Where FraseH.Codigo = " + "'" + ZImpreI(Ciclo) + "'"
        Rem         spFraseH = Sql1 + Sql2 + Sql3
        Rem         Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        Rem         If rstFraseH.RecordCount > 0 Then
        Rem             ZImpreI(Ciclo) = rstFraseH!Descripcion
        Rem             rstFraseH.Close
        Rem         End If
        Rem     End If
        Rem Next Ciclo
    
        Call Conecta_Empresa
        
        Erase ZZImpreFrase
        LugarFrase = 1
        
        ZZCorte = 185
        
        For Ciclo = 1 To 99
        
            If Trim(ZImpreI(Ciclo)) <> "" Then
            
                AA1 = Trim(ZImpreI(Ciclo))
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreI(Ciclo)) + " "
                
                Do
                
                    ZZHastaIII = Len(ZZImpreFrase(LugarFrase))
                    
                    ZZHastaII = Len(ZZImpreFrase(LugarFrase))
                    For Da = 1 To ZZHastaIII
                        If Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) >= 65 And Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) <= 90 Then
                            ZZHastaII = ZZHastaII + 0.5
                        End If
                    Next Da
                
                    If ZZHastaII > ZZCorte Then
                    
                        For Da = ZZHastaIII - 1 To 1 Step -1
                            If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "-" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "+" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "," Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "/" Then

                                Auxi = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                ZZHastaIII = Len(Auxi)
                                ZZHastaII = 0
                                For DaIII = 1 To ZZHastaIII
                                    ZZHastaII = ZZHastaII + 1
                                    If Asc(Mid$(Auxi, DaIII, 1)) >= 65 And Asc(Mid$(Auxi, DaIII, 1)) <= 90 Then
                                        ZZHastaII = ZZHastaII + 0.5
                                    End If
                                Next DaIII
                                If ZZHastaII <= ZZCorte Then
                                    Auxi = ZZImpreFrase(LugarFrase)
                                    ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                    LugarFrase = LugarFrase + 1
                                    ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Mid$(Auxi, Da + 1, ZZCorte)
                                    Exit For
                                End If
                                
                            End If
                        Next Da

                            Else
                            
                        Exit Do
                        
                    End If
                Loop
            End If
        
        Next Ciclo
            
        
        For Ciclo = 1 To 99
        
            If Trim(ZImpreII(Ciclo)) <> "" Then
            
                AA1 = Trim(ZImpreII(Ciclo))
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreII(Ciclo)) + " "
                
                Do
                
                    ZZHastaIII = Len(ZZImpreFrase(LugarFrase))
                    
                    ZZHastaII = Len(ZZImpreFrase(LugarFrase))
                    For Da = 1 To ZZHastaIII
                        If Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) >= 65 And Asc(Mid$(ZZImpreFrase(LugarFrase), Da, 1)) <= 90 Then
                            ZZHastaII = ZZHastaII + 0.5
                        End If
                    Next Da
                
                    If ZZHastaII > ZZCorte Then
                    
                        For Da = ZZHastaIII - 1 To 1 Step -1
                            If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "-" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "+" Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "," Or Mid$(ZZImpreFrase(LugarFrase), Da, 1) = "/" Then
                            
                                Auxi = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                ZZHastaIII = Len(Auxi)
                                ZZHastaII = 0
                                For DaIII = 1 To ZZHastaIII
                                    ZZHastaII = ZZHastaII + 1
                                    If Asc(Mid$(Auxi, DaIII, 1)) >= 65 And Asc(Mid$(Auxi, DaIII, 1)) <= 90 Then
                                        ZZHastaII = ZZHastaII + 0.5
                                    End If
                                Next DaIII
                                If ZZHastaII <= ZZCorte Then
                                    Auxi = ZZImpreFrase(LugarFrase)
                                    ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                                    LugarFrase = LugarFrase + 1
                                    ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Mid$(Auxi, Da + 1, ZZCorte)
                                    Exit For
                                End If
                                
                            End If
                        Next Da

                            Else
                            
                        Exit Do
                        
                    End If
                Loop
            End If
        
        Next Ciclo
        
        
        
        
        
            
        For Ciclo = 1 To 99
        
            If Trim(ZImpreIII(Ciclo)) <> "" Then
            
                ZZImpreFrase(LugarFrase) = ZZImpreFrase(LugarFrase) + Trim(ZImpreIII(Ciclo)) + " "
                
                If Len(ZZImpreFrase(LugarFrase)) > ZZCorte Then
                    
                    For Da = ZZCorte To 1 Step -1
                        If Mid$(ZZImpreFrase(LugarFrase), Da, 1) = Space$(1) Then
                            Auxi = ZZImpreFrase(LugarFrase)
                            ZZImpreFrase(LugarFrase) = Mid$(ZZImpreFrase(LugarFrase), 1, Da)
                            LugarFrase = LugarFrase + 1
                            ZZImpreFrase(LugarFrase) = Mid$(Auxi, Da + 1, ZZCorte)
                            Exit For
                        End If
                    Next Da
                    
                End If
            End If
        
        Next Ciclo
            
            
            
            
            
            
        For Ciclo = 1 To 19
        
            If Trim(ZZImpreFrase(Ciclo)) <> "" Then
            
                For CicloII = 1 To ZZHasta
                
                    If Mid$(ZZImpreFrase(Ciclo), CicloII, 1) = Space$(1) Then
                        ZZImpreFrase(Ciclo) = Mid$(ZZImpreFrase(Ciclo), 1, CicloII) + " " + Mid$(ZZImpreFrase(Ciclo), CicloII + 1, ZZCorte)
                        ZZHasta = Len(Trim(ZZImpreFrase(Ciclo)))
                        CicloII = CicloII + 1
                        If CicloII = ZZCorte Or ZZHasta = ZZCorte Then
                            Exit For
                        End If
                    End If
                    
                Next CicloII
                
                ZZImpreFrase(Ciclo) = Trim(ZZImpreFrase(Ciclo))
                
                
            End If
            
        Next Ciclo
                        
                    
            
        
            
            
            
            
            
            
            
            
    
    
    
    


    With rstEtiquetaII
        For Da = 1 To ZCantidad
            .Index = "Codigo"
            .AddNew
    
            WLote = Lote.Text
            Call Ceros(WLote, 6)
    
            WCantidad = Kilos.Text
            Call Ceros(WCantidad, 4)
    
            ZDa = Int((Da - 1) / 2)
    
            !Codigo = Da
            !Terminado = Codigo.Text
            !Lote = WLote
            !Cliente = ""
            !Cantidad = Val(Kilos.Text)
            !Nombre = "Fec.Lau.: " + Fecha.Text
            If ZVencimiento <> "00/00/0000" Then
                !Impre1 = "Fec.Rea.:" + ZVencimiento
                    Else
                !Impre1 = ""
            End If
            !Conservacion = !Impre1
            
            Rem XTipoPro = ""
            Rem If Left$(WTerminado, 2) = "PT" Then
            Rem     If XCodigo >= 0 And XCodigo <= 999 Then
            Rem         XTipoPro = "CO"
            Rem             Else
            Rem         If XCodigo >= 11000 And XCodigo <= 12999 Then
            Rem             XTipoPro = "CO"
            Rem         End If
            Rem     End If
            Rem End If
            Rem If XTipoPro = "CO" Then
            Rem     !Nombre = ""
            Rem     !Impre1 = ""
            Rem     !Conservacion = !Impre1
            Rem End If
            
            
            
            !Razon = "L:" + Partida.Text
            !DirEntrega = Kilos.Text + " Kgs."
            !Clase = WClase
            !Intervencion = WIntervencion
            !Naciones = WNaciones
            !Embalaje = WEmbalaje
            !Bruto = 0
            If Da = ZCantidad And ZMarca = 1 Then
                !Bruto = 1
            End If
            !Neto = ZDa
            If Val(Wempresa) = 1 Or Val(Wempresa) = 3 Or Val(Wempresa) = 5 Or Val(Wempresa) = 6 Or Val(Wempresa) = 7 Or Val(Wempresa) = 10 Or Val(Wempresa) = 11 Then
                !Observaciones = "CONTROL de CALIDAD"
                    Else
                !Observaciones = "C.C.   PELLITAL"
            End If
            !ConservacionII = DesCodigo.Caption
            
            
            !foto1 = 0
            !foto2 = 0
            !foto3 = 0
            !foto4 = 0
            !foto5 = 0
            
            For Ciclo = 1 To 9
                If ZZLogo(Ciclo) <> 0 Then
                    Select Case ZZLogo(Ciclo)
                        Case 1
                            !foto1 = Ciclo
                        Case 2
                            !foto2 = Ciclo
                        Case 3
                            !foto3 = Ciclo
                        Case 4
                            !foto4 = Ciclo
                        Case 5
                            !foto5 = Ciclo
                        Case Else
                    End Select
                End If
            Next Ciclo
            
            
            .Update
        Next Da
    End With
    


    With rstEtiquetaIII
        For Da = 1 To ZCantidad
            .Index = "Codigo"
            .AddNew
            !Codigo = Da
            
            aa = Len(ZZImpreFrase(1))
            
            !Frase1 = ZZImpreFrase(1)
            !Frase2 = ZZImpreFrase(2)
            !Frase3 = ZZImpreFrase(3)
            !Frase4 = ZZImpreFrase(4)
            !Frase5 = ZZImpreFrase(5)
            !Frase6 = ZZImpreFrase(6)
            !Frase7 = ZZImpreFrase(7)
            !Frase8 = ZZImpreFrase(8)
            !Frase9 = ZZImpreFrase(9)
            !Frase10 = ZZImpreFrase(10)
            
            
            .Update
        Next Da
    End With
    
    
    
    
    
    
    
        
        
        
    With rstEtiquetaIV
        For Da = 1 To ZCantidad
            .Index = "Codigo"
            .AddNew
            !Codigo = Da
            
            !Frase20 = ""
            If ZZPalabra = 1 Then
                !Frase20 = "PELIGRO"
            End If
            If ZZPalabra = 2 Then
                !Frase20 = "ATENCION"
            End If
            
            !Frase11 = ZZImpreFrase(11)
            !Frase12 = ZZImpreFrase(12)
            !Frase13 = ZZImpreFrase(13)
            !Frase14 = ZZImpreFrase(14)
            !Frase15 = ZZImpreFrase(15)
            !Frase16 = ZZImpreFrase(16)
            !Frase17 = ZZImpreFrase(17)
            !Frase18 = ZZImpreFrase(18)
            !Frase19 = ZZImpreFrase(19)
            
            !Frase20 = ""
            If ZZPalabra = 1 Then
                !Frase20 = "PELIGRO"
            End If
            If ZZPalabra = 2 Then
                !Frase20 = "ATENCION"
            End If
            
            
            .Update
        Next Da
    End With
        
        
        
        
        
        
        
        
    
    

    Listado.WindowTitle = "Emision de Etiquetas"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    Select Case Val(Wempresa)
        Case 2, 4, 8, 9
            Listado.ReportFileName = "Etinuevanormapelli.rpt"
       Case Else
            Listado.ReportFileName = "Etinuevanormamp.rpt"
    End Select
    
    Rem Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"

    Listado.Destination = 1
     Listado.Destination = 0
    Listado.PrinterCopies = 1
    Listado.Action = 1
                
    
    Da = 0
    With rstEtiquetaII
        .Index = "Codigo"
        .Seek ">=", Da
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
    
    
    Lote.Text = ""
    Codigo.Text = ""
    DesCodigo.Caption = ""
    Fecha.Text = ""
    Kilos.Text = ""
    Cantidad.Text = ""
    TipoEtiqueta.ListIndex = 0
    
    Partida.Visible = True
    Partida.Text = ""
             
    Informe.Visible = True
    Informe.Text = ""
    Lote.Enabled = True
    Lote.Visible = True
    Codigo.Locked = True
    Fecha.Visible = True
    Lote.SetFocus
    
    Exit Sub

WError:

    Resume Next

End Sub

Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem "M.P."
    Tipo.AddItem "P.T."
    
    Tipo.ListIndex = 0
    
    TipoEtiqueta.Clear
    
    TipoEtiqueta.AddItem "Etiqueta Verde"
    TipoEtiqueta.AddItem "Etiqueta Verde Farma"
    TipoEtiqueta.AddItem "Etiqueta Colorante"
    TipoEtiqueta.AddItem "Etiqueta NI"
    TipoEtiqueta.ListIndex = 0

    Lote.Text = ""
    Codigo.Text = ""
    DesCodigo.Caption = ""
    Fecha.Text = ""
    Kilos.Text = ""
    Cantidad.Text = ""
    
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        If Tipo.ListIndex = 0 Then
    
            Ingresa = "N"
            
            XEmpresa = Wempresa
            
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
                    ZHasta = 7
                Case Else
                    Empe(1, 1) = "0002"
                    Empe(1, 2) = "Empresa02"
                    Empe(2, 1) = "0004"
                    Empe(2, 2) = "Empresa04"
                    Empe(3, 1) = "0008"
                    Empe(3, 2) = "Empresa08"
                    Empe(4, 1) = "0009"
                    Empe(4, 2) = "Empresa09"
                    ZHasta = 4
            End Select
    
            For A = 1 To ZHasta
            
                Wempresa = Empe(A, 1)
                txtOdbc = Empe(A, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Laudo"
                ZSql = ZSql + " Where Laudo.Laudo = " + "'" + Lote.Text + "'"
                spLaudo = ZSql
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                
                    Codigo.Text = rstLaudo!Articulo
                    ZArticulo = Left$(Codigo.Text, 2)
                    If ZArticulo = "DS" Then
                        ZArticuloII = Mid$(Codigo.Text, 4, 3)
                        If Val(ZArticuloII) < 100 Then
                            ZArticulo = ""
                        End If
                    End If
                    
                    ZZPartiOri = rstLaudo!PartiOri
                    ZZLote = Lote.Text
                    
                    If ZArticulo = "DY" Or ZArticulo = "DS" Then
                        Partida.Text = rstLaudo!PartiOri
                            Else
                        Partida.Text = Lote.Text
                    End If
                    
                    Informe.Text = rstLaudo!Informe
                    Fecha.Text = rstLaudo!Fecha
                    ZVencimiento = IIf(IsNull(rstLaudo!FechaVencimiento), "00/00/0000", rstLaudo!FechaVencimiento)
                    Ingresa = "S"
                    rstLaudo.Close
                    
                    Sql1 = "Select *"
                    Sql2 = " FROM PrueArt"
                    Sql3 = " WHERE Lote = " + "'" + ZZLote + "'"
                    spPrueart = Sql1 + Sql2 + Sql3
                    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
                    If rstPrueart.RecordCount > 0 Then
                        Partida.Text = ZZLote
                        rstPrueart.Close
                            Else
                        Partida.Text = ZZLote
                    End If
                    
                End If
                
                
            Next A
            
            Call Conecta_Empresa
            
            If Ingresa = "N" Then
            
                Lote.SetFocus
                
                    Else
                
                spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    DesCodigo.Caption = rstArticulo!Descripcion
                    WImpreadi = ""
                    WClase = ""
                    WIntervencion = ""
                    WNaciones = ""
                    WEmbalaje = ""
                    WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
                    WIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
                    WNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
                    WEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
                    rstArticulo.Close
                End If
                
                Kilos.SetFocus
                
            End If
            
                    Else
            
            Ingresa = "N"
            
            XEmpresa = Wempresa
            
            Select Case Val(Wempresa)
                Case 1, 3, 5, 6, 7, 10, 11
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
                    ZHasta = 7
                Case Else
                    Empe(1, 1) = "0002"
                    Empe(1, 2) = "Empresa02"
                    Empe(2, 1) = "0004"
                    Empe(2, 2) = "Empresa04"
                    Empe(3, 1) = "0008"
                    Empe(3, 2) = "Empresa08"
                    Empe(4, 1) = "0009"
                    Empe(4, 2) = "Empresa09"
                    ZHasta = 4
            End Select
    
            For A = 1 To ZHasta
            
                Wempresa = Empe(A, 1)
                txtOdbc = Empe(A, 2)
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM Hoja"
                ZSql = ZSql + " Where Hoja.Hoja = " + "'" + Lote.Text + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                If rstHoja.RecordCount > 0 Then
                    Codigo.Text = rstHoja!Producto
                    Partida.Text = Lote.Text
                    Informe.Text = ""
                    ZVencimiento = "00/00/0000"
                    Fecha.Text = rstHoja!Fecha
                    Ingresa = "S"
                    rstHoja.Close
                    Exit For
                End If
                
            Next A
            
            Call Conecta_Empresa
            
            If Ingresa = "N" Then
            
                Lote.SetFocus
                
                    Else
                
                spTerminado = "ConsultaTerminado " + "'" + Codigo.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    DesCodigo.Caption = rstTerminado!Descripcion
                    WImpreadi = ""
                    WClase = ""
                    WIntervencion = ""
                    WNaciones = ""
                    WEmbalaje = ""
                    WImpreadi = IIf(IsNull(rstTerminado!Impreadi), "", rstTerminado!Impreadi)
                    WClase = IIf(IsNull(rstTerminado!Clase), "", rstTerminado!Clase)
                    WIntervencion = IIf(IsNull(rstTerminado!Intervencion), "", rstTerminado!Intervencion)
                    WNaciones = IIf(IsNull(rstTerminado!Naciones), "", rstTerminado!Naciones)
                    WEmbalaje = IIf(IsNull(rstTerminado!Embalaje), "", rstTerminado!Embalaje)
                    rstTerminado.Close
                End If
                
                Kilos.SetFocus
                
            End If
                
        End If
        
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Kilos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cantidad.SetFocus
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Kilos.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
    OPEN_FILE_EtiquetaII
    OPEN_FILE_EtiquetaIII
    OPEN_FILE_EtiquetaIV
End Sub

Private Sub TipoEtiqueta_Click()

     If TipoEtiqueta.ListIndex = 3 Then
        Lote.Text = ""
        Codigo.Text = ""
        DesCodigo.Caption = ""
        Fecha.Text = ""
        Kilos.Text = ""
        Cantidad.Text = ""
        Partida.Visible = False
        Informe.Visible = False
        Fecha.Visible = False
        Lote.Visible = False
        Codigo.Enabled = True
        Codigo.Locked = False
        Codigo.SetFocus
        Tipo.ListIndex = 0
            Else
        Partida.Visible = True
        Rem Partida.Text = ""
        Informe.Visible = True
        Rem Informe.Text = ""
        Lote.Enabled = True
        Lote.Visible = True
        Rem Lote.Text = ""
        Rem Kilos.Text = ""
        Rem Cantidad.Text = ""
        Codigo.Enabled = True
        Rem Codigo.Text = ""
        Fecha.Visible = True
        Rem Fecha.Text = ""
    End If

End Sub
