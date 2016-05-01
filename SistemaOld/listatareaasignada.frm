VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgListaTareaAsignada 
   Caption         =   "Listado de Tareas Asignadas a Responsables"
   ClientHeight    =   4245
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4245
   ScaleWidth      =   8145
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox TipoIII 
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
         Left            =   2040
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox TipoII 
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
         Left            =   2040
         TabIndex        =   3
         Top             =   1680
         Width           =   2295
      End
      Begin VB.ComboBox TipoI 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   1200
         Width           =   2295
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
         Left            =   2880
         TabIndex        =   10
         Top             =   3000
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
         Left            =   1440
         TabIndex        =   9
         Top             =   3000
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
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   8
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
         Left            =   4440
         MaskColor       =   &H00000000&
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Hasta 
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Top             =   600
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
         Left            =   2040
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label Label5 
         Caption         =   "Estado"
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
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
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
         Left            =   360
         TabIndex        =   14
         Top             =   600
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
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Asignado"
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
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Emisor"
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
         Left            =   360
         TabIndex        =   11
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7560
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgListaTareaAsignada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim ZResponsable(1000) As String
Dim ZLugar As Integer




Private Sub Acepta_Click()

    Rem On Error GoTo WError
    
    Listado.WindowTitle = "Listado de Tareas Asignadas a Responsables"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height
    
    If Desde.Text = "  /  /    " Then
        Desde.Text = "01/01/2000"
    End If
    If Hasta.Text = "  /  /    " Then
        Hasta.Text = "01/01/2999"
    End If
    
    DesdeFecha = Right$(Desde.Text, 4) + Mid$(Desde.Text, 4, 2) + Left$(Desde.Text, 2)
    HastaFecha = Right$(Hasta.Text, 4) + Mid$(Hasta.Text, 4, 2) + Left$(Hasta.Text, 2)
    
    If TipoI.ListIndex = 0 Then
        ZDesdeI = "0"
        ZHastaI = "9999"
            Else
        ZDesdeI = ZResponsable(TipoI.ListIndex)
        ZHastaI = ZResponsable(TipoI.ListIndex)
    End If
    
    If TipoII.ListIndex = 0 Then
        ZDesdeII = "0"
        ZHastaII = "9999"
            Else
        ZDesdeII = ZResponsable(TipoII.ListIndex)
        ZHastaII = ZResponsable(TipoII.ListIndex)
    End If
            
    If TipoIII.ListIndex = 0 Then
        ZDesdeIII = "1"
        ZHastaIII = "1"
            Else
        ZDesdeIII = "0"
        ZHastaIII = "9999"
    End If
    
    
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Planifica.Responsable, Planifica.Numero, Planifica.Fecha, Planifica.Vencimiento, Planifica.OrdVencimiento, Planifica.ResponsableII, Planifica.Estado, Planifica.Observaciones, Planifica.DesResponsable, Planifica.DesResponsableII, Planifica.Descripcion " _
            + "From  " _
            + DSQ + ".dbo.Planifica Planifica " _
            + "Where " _
            + "Planifica.Responsable >= " + ZDesdeI + " AND " _
            + "Planifica.Responsable <= " + ZHastaI + " AND " _
            + "Planifica.OrdVencimiento >= '" + DesdeFecha + "' AND " _
            + "Planifica.OrdVencimiento <= '" + HastaFecha + "' AND " _
            + "Planifica.ResponsableII >= " + ZDesdeII + " AND " _
            + "Planifica.ResponsableII <= " + ZHastaII + " AND " _
            + "Planifica.Estado >= " + ZDesdeIII + " AND " _
            + "Planifica.Estado <= " + ZHastaIII
    
    Uno = "{Planifica.ResponsableII} in " + ZDesdeII + " to " + ZHastaII
    Dos = " and {Planifica.OrdVencimiento} in " + Chr$(34) + DesdeFecha + Chr$(34) + " to " + Chr$(34) + HastaFecha + Chr$(34)
    Tres = " and {Planifica.Estado} in " + ZDesdeIII + " to " + ZHastaIII
    Cuatro = " and {Planifica.Responsable} in " + ZDesdeI + " to " + ZHastaI
    
    Rem Uno = ""
    Rem Dos = ""
    Rem Tres = ""
    Rem Cuatro = ""
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
    Listado.SelectionFormula = Uno + Dos + Tres + Cuatro
    
    Listado.Connect = Connect()
    
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.ReportFileName = "ListaTareaAsignada.Rpt"
    Listado.Action = 1
    
    Exit Sub

WError:
    Resume Next
    
End Sub

Private Sub Cancela_click()
    PrgListaTareaAsignada.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()
    
    TipoI.Clear
    TipoII.Clear
    
    TipoI.AddItem "Total"
    TipoII.AddItem "Total"
    
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM ResponsableSac"
    spResponsableSac = ZSql
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        With rstResponsableSac
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    ZResponsable(ZLugar) = rstResponsableSac!Codigo
                    
                    If ZZOperadorResponsable = rstResponsableSac!Codigo Then
                        ZPuntero = ZLugar
                    End If
                
                    TipoI.AddItem Trim(rstResponsableSac!Descripcion)
                    TipoII.AddItem Trim(rstResponsableSac!Descripcion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstResponsableSac.Close
    End If
    
    TipoI.ListIndex = ZPuntero
    TipoII.ListIndex = 0
    
    TipoIII.Clear
    
    TipoIII.AddItem "Pendiente"
    TipoIII.AddItem "Total"
    
    TipoIII.ListIndex = 0
    
    Desde.Text = "  /  /    "
    Hasta.Text = "  /  /    "
    
    Panta.Value = True
    Impresora.Value = False
    
    Frame2.Visible = True
    
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
        Desde.SetFocus
    End If
End Sub

