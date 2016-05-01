VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEtiContra 
   Caption         =   "Impresion de Etiquetas de Muestra Simple"
   ClientHeight    =   5685
   ClientLeft      =   1170
   ClientTop       =   1485
   ClientWidth     =   10425
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   10425
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   9855
      Begin VB.TextBox Analista 
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
         MaxLength       =   20
         TabIndex        =   19
         Text            =   "  "
         Top             =   2880
         Width           =   2295
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
         TabIndex        =   17
         Text            =   "  "
         Top             =   1440
         Width           =   2055
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
         MaxLength       =   20
         TabIndex        =   15
         Text            =   "  "
         Top             =   1440
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
         TabIndex        =   14
         Text            =   " "
         Top             =   1920
         Width           =   1455
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
         TabIndex        =   12
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
         TabIndex        =   10
         Text            =   " "
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
      Begin MSMask.MaskEdBox Codigo 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
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
      Begin VB.Label Label2 
         Caption         =   "Analista"
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
         Top             =   2880
         Width           =   2055
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
         TabIndex        =   18
         Top             =   1440
         Width           =   855
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
         TabIndex        =   16
         Top             =   1440
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
         TabIndex        =   13
         Top             =   1920
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
         Left            =   3840
         TabIndex        =   11
         Top             =   960
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
         TabIndex        =   9
         Top             =   2400
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
         Width           =   2055
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
         Top             =   960
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
Attribute VB_Name = "PrgEtiContra"
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
    
        ZCantidad = Int(Val(Cantidad.Text) / 2)
        If ZCantidad * 2 <> Val(Cantidad.Text) Then
            ZCantidad = ZCantidad + 1
        End If
        
        ZLugar = 0
        
        With rstEtiqueta
            For Da = 1 To ZCantidad
                .Index = "Codigo"
                .AddNew
                
                WLote = Lote.Text
                Call Ceros(WLote, 6)
                
                WCantidad = "0"
                Call Ceros(WCantidad, 4)
                
                ZDa = Da
                
                !Codigo = ZDa
                !Terminado = Codigo.Text
                !razon = DesCodigo.Caption
                !Lote = Val(Informe.Text)
                ZLugar = ZLugar + 1
                !Impre1 = Str$(ZLugar) + "/" + Str$(ZCantidad * 2)
                !Observaciones = Analista.Text
                !Nombre = "Fecha : " + Fecha.Text
                ZLugar = ZLugar + 1
                !Cliente = ""
                !Cantidad = 0
                !DirEntrega = Str$(ZLugar) + "/" + Str$(ZCantidad * 2)
                !Clase = WClase
                !Intervencion = WIntervencion
                !Naciones = WNaciones
                !Embalaje = WEmbalaje
                !Bruto = 0
                !Neto = ZDa
                .Update
            Next Da
        End With

        Listado.WindowTitle = "Emision de Etiquetas"
        Listado.WindowTop = 0
        Listado.WindowLeft = 0
        Listado.WindowWidth = Screen.Width
        Listado.WindowHeight = Screen.Height
        
        If Len(Trim(DesCodigo.Caption)) < 23 Then
            Listado.ReportFileName = "WEtiContra.rpt"
                Else
            Listado.ReportFileName = "WEtiContraii.rpt"
        End If
    
        Listado.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
        Listado.Destination = 1
        Listado.PrinterCopies = 1
        Listado.Action = 1
    
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
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Cancela_click()

    With rstEmpresa
        .Close
    End With
    PrgEtiContra.Hide
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

Sub Form_Load()

    Lote.Text = ""
    Codigo.Text = "  -   -   "
    DesCodigo.Caption = ""
    Partida.Text = ""
    Informe.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Analista.Text = ""
    Cantidad.Text = ""
    
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)

    On Error GoTo WError

    If KeyAscii = 13 Then
    
        If Val(Lote.Text) = 0 Then
            Codigo.SetFocus
            Exit Sub
        End If
    
        Ingresa = "N"
            
        XEmpresa = WEmpresa
            
        Select Case Val(WEmpresa)
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
            
            WEmpresa = Empe(A, 1)
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
                If Left$(Codigo.Text, 2) = "DY" Or Left$(Codigo.Text, 2) = "DW" Or Left$(Codigo.Text, 2) = "DS" Then
                    Partida.Text = rstLaudo!partiori
                        Else
                    Partida.Text = Lote.Text
                End If
                Informe.Text = rstLaudo!Informe
                Fecha.Text = rstLaudo!Fecha
                ZVencimiento = IIf(IsNull(rstLaudo!fechavencimiento), "00/00/0000", rstLaudo!fechavencimiento)
                Ingresa = "S"
                rstLaudo.Close
                Exit For
            End If
                
        Next A
            
        Call Conecta_Empresa
            
        If Ingresa = "N" Then
            
            Codigo.SetFocus
                
                Else
                
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesCodigo.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
                
            Cantidad.SetFocus
                
        End If
        
    End If
    
    Exit Sub

WError:

    Resume Next
    
End Sub


Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesCodigo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
            Cantidad.SetFocus
                Else
            Codigo.SetFocus
        End If
    End If
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Analista.SetFocus
    End If
End Sub

Private Sub Analista_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rem Cantidad.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    OPEN_FILE_Etiqueta
End Sub

