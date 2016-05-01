VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgAgendaPlanificaI 
   AutoRedraw      =   -1  'True
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   11880
   Begin VB.Frame PantaModifica 
      Height          =   3375
      Left            =   1680
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox clavemodifica 
         Height          =   285
         Left            =   840
         TabIndex        =   19
         Top             =   2520
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton GrabaModifica 
         Caption         =   "Graba"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3600
         TabIndex        =   18
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox Observaciones 
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
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   16
         Top             =   1800
         Width           =   6255
      End
      Begin VB.ComboBox Estado 
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox Descripcion 
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1080
         Width           =   6255
      End
      Begin VB.TextBox Responsable 
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox ResponsableII 
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Descripcion"
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
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Responsables"
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Descripcion"
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label DesResponsable 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Responsable Emisor"
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
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable Destinatario"
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
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label DesResponsableII 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3720
         TabIndex        =   8
         Top             =   720
         Width           =   3495
      End
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
      Height          =   735
      Left            =   4320
      TabIndex        =   4
      Top             =   7080
      Width           =   1215
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7680
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   6000
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   6855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgAgendaPlanificaI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPlanifica As Recordset
Dim spPlanifica As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim XParam As String
Dim ZZLugar As Integer

Dim ZZAyuda(1000) As String

Private Sub Cancela_click()
    PrgAgendaPlanificaI.Hide
    Unload Me
    PrgAgendaTotal.Show
End Sub

Sub Form_Load()

    Call Limpia_Ayuda
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ResponsableSac"
    ZSql = ZSql + " Order by ResponsableSac.Codigo"
    spResponsableSac = ZSql
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        With rstResponsableSac
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZAyuda(rstResponsableSac!Codigo) = Trim(rstResponsableSac!Descripcion)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstResponsableSac.Close
    End If
    
    


    ZLugar = 0

    For Ciclo = 1 To 100
    
        If Val(ZZPasaDatos(Ciclo, 1)) <> 0 Then
        
            ZClave = ZZPasaDatos(Ciclo, 1)
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Planifica"
            ZSql = ZSql + " Where Planifica.Clave = " + "'" + ZClave + "'"
            spPlanifica = ZSql
            Set rstPlanifica = db.OpenRecordset(spPlanifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstPlanifica.RecordCount > 0 Then
            
                ZLugar = ZLugar + 1
                
                Pantalla.TextMatrix(ZLugar, 1) = ZZAyuda(rstPlanifica!Responsable)
                Pantalla.TextMatrix(ZLugar, 2) = rstPlanifica!Descripcion
                Pantalla.TextMatrix(ZLugar, 3) = rstPlanifica!Vencimiento
                Pantalla.TextMatrix(ZLugar, 4) = rstPlanifica!Observaciones
                Pantalla.TextMatrix(ZLugar, 5) = rstPlanifica!Clave
            
                rstPlanifica.Close
                
            End If
            
        End If
        
    Next Ciclo
    
End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    Select Case ZZLugar
        Case 1, 2, 4
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
            
            Pantalla.ColWidth(0) = 200
            Pantalla.Row = 0
            
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Codigo"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 6000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
            Rem DESPILEGA LOS TITULOS
            
            WTitulo(1).Visible = False
            WTitulo(2).Visible = False
            
        Case Else
            Pantalla.FixedCols = 1
            Pantalla.Cols = 6
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
            
            Pantalla.ColWidth(0) = 200
            Pantalla.Row = 0
            
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Emisor"
                        Pantalla.ColWidth(Ciclo) = 2000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 2
                        Pantalla.Text = "Descripcion"
                        Pantalla.ColWidth(Ciclo) = 5000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 3
                        Pantalla.Text = "Vto."
                        Pantalla.ColWidth(Ciclo) = 1200
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 4
                        Pantalla.Text = "Observaciones"
                        Pantalla.ColWidth(Ciclo) = 2000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 5
                        Pantalla.Text = ""
                        Pantalla.ColWidth(Ciclo) = 50
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
            Rem DESPILEGA LOS TITULOS
            
            WTitulo(1).Visible = False
            WTitulo(2).Visible = False
            
    End Select
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        Rem WTitulo(Ciclo).Text = Pantalla.Text
        Rem WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        Rem WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        Rem WTitulo(Ciclo).Width = Pantalla.CellWidth
        Rem WTitulo(Ciclo).Height = Pantalla.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Rem Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub


Private Sub GrabaModifica_Click()
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Planifica SET "
    ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
    ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "'"
    ZSql = ZSql + " Where Clave = " + "'" + clavemodifica.Text + "'"
    spPlanifica = ZSql
    Set rstPlanifica = db.OpenRecordset(spPlanifica, dbOpenSnapshot, dbSQLPassThrough)
    
    PantaModifica.Visible = False
    
    Pantalla.TextMatrix(Pantalla.Row, 4) = Observaciones.Text
    

End Sub

Private Sub Pantalla_Click()

    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "Pendiente"
    Estado.AddItem "Finalizado"
    
    Estado.ListIndex = 0
    
    ZClave = Pantalla.TextMatrix(Pantalla.Row, 5)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Planifica"
    ZSql = ZSql + " Where Planifica.Clave = " + "'" + ZClave + "'"
    spPlanifica = ZSql
    Set rstPlanifica = db.OpenRecordset(spPlanifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstPlanifica.RecordCount > 0 Then
    
        ZLugar = ZLugar + 1
        
        Responsable.Text = rstPlanifica!Responsable
        ResponsableII.Text = rstPlanifica!ResponsableII
        DesResponsable.Caption = ZZAyuda(rstPlanifica!Responsable)
        DesResponsableII.Caption = ZZAyuda(rstPlanifica!ResponsableII)
        Descripcion.Text = Trim(rstPlanifica!Descripcion)
        Estado.ListIndex = rstPlanifica!Estado
        Observaciones.Text = Trim(rstPlanifica!Observaciones)
        clavemodifica.Text = ZClave
    
        rstPlanifica.Close
        
    End If
    
    PantaModifica.Visible = True
    
    Observaciones.SetFocus
    
End Sub
