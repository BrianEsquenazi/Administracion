VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgEquiposFabrica 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Equipos, Control y Instrucciones de Seguridad"
   ClientHeight    =   6945
   ClientLeft      =   300
   ClientTop       =   1005
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   6945
   ScaleWidth      =   11430
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox DescripcionIII 
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
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   14
      Top             =   1800
      Width           =   9015
   End
   Begin VB.TextBox DescripcionII 
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
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   13
      Top             =   1440
      Width           =   9015
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Hasta 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   11
         Text            =   " "
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Desde 
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
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   10
         Text            =   " "
         Top             =   360
         Width           =   855
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
         TabIndex        =   9
         Top             =   1200
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
         TabIndex        =   8
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "equiposFabrica.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "equiposFabrica.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "equiposFabrica.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "equiposFabrica.frx":0A56
         ToolTipText     =   "Cancela la Impresion"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Codigo"
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
         Left            =   720
         TabIndex        =   7
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
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
         Left            =   720
         TabIndex        =   6
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5280
      TabIndex        =   12
      Top             =   2280
      Width           =   2055
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   240
         MouseIcon       =   "equiposFabrica.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "equiposFabrica.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1080
         MouseIcon       =   "equiposFabrica.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "equiposFabrica.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.TextBox Codigo 
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
      Left            =   2160
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   600
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7440
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WCampaña.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Bancos"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   975
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
      Left            =   2160
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1080
      Width           =   9015
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "equiposFabrica.frx":1D30
      MousePointer    =   99  'Custom
      Picture         =   "equiposFabrica.frx":203A
      ToolTipText     =   "Impresion "
      Top             =   2520
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "equiposFabrica.frx":287C
      MousePointer    =   99  'Custom
      Picture         =   "equiposFabrica.frx":2B86
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "equiposFabrica.frx":33C8
      MousePointer    =   99  'Custom
      Picture         =   "equiposFabrica.frx":36D2
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "equiposFabrica.frx":3F14
      MousePointer    =   99  'Custom
      Picture         =   "equiposFabrica.frx":421E
      ToolTipText     =   "Elimina el Registro"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "equiposFabrica.frx":4A60
      MousePointer    =   99  'Custom
      Picture         =   "equiposFabrica.frx":4D6A
      ToolTipText     =   "Salida"
      Top             =   2520
      Width           =   480
   End
   Begin VB.Label lblLabels 
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
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo "
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
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   660
      Width           =   2295
   End
End
Attribute VB_Name = "PrgEquiposFabrica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEquipoFabrica As Recordset
Dim spEquipoFabrica As String

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        Descripcion.Text = Trim(rstEquipoFabrica!Descripcion)
        DescripcionII.Text = Trim(rstEquipoFabrica!DescripcionII)
        DescripcionIII.Text = Trim(rstEquipoFabrica!DescripcionIII)
        rstEquipoFabrica.Close
    End If
    
End Sub

Private Sub Acepta_Click()
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = ""
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{EquipoFabrica.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            rstEquipoFabrica.Close
            Sql1 = "UPDATE EquipoFabrica SET "
            Sql2 = " Descripcion = " + "'" + Descripcion.Text + "',"
            Sql3 = " DescripcionII = " + "'" + DescripcionII.Text + "',"
            Sql4 = " DescripcionIII = " + "'" + DescripcionIII.Text + "'"
            Sql5 = " Where Codigo = " + "'" + Codigo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
                Else
            Sql1 = "INSERT INTO EquipoFabrica ("
            Sql2 = "Codigo ,"
            Sql3 = "Descripcion ,"
            Sql4 = "DescripcionII ,"
            Sql5 = "DescripcionIII )"
            Sql6 = "Values ("
            Sql7 = "'" + Codigo.Text + "',"
            Sql8 = "'" + Descripcion.Text + "',"
            Sql9 = "'" + DescripcionII.Text + "',"
            Sql10 = "'" + DescripcionIII.Text + "')"
            spEquipoFabrica = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM EquipoFabrica"
        Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
        spEquipoFabrica = Sql1 + Sql2 + Sql3
        Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipoFabrica.RecordCount > 0 Then
            rstEquipoFabrica.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE EquipoFabrica"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spEquipoFabrica = Sql1 + Sql2
                Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""

    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgEquiposFabrica.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Command1_Click()
    
    ZZNumero = "3054916508319000860353368223354201009108"
    
    lcstart = Chr(40)
    lcstop = Chr(41)
    barralargo = ZZNumero
    
    For lni = 1 To Len(barralargo) Step 2
    If Val(Mid(barralargo, lni, 2)) < 50 Then
    lccar = lccar + Chr(Val(Mid(barralargo, lni, 2)) + 48)
    Else
    lccar = lccar + Chr(Val(Mid(barralargo, lni, 2)) + 142)
    End If
    Next
    barralargo = lccar
    
    DescripcionII.Text = barralargo

End Sub

Private Sub Lista_Click()
    Desde.Text = "  -     -   "
    Hasta.Text = "  -     -   "
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub DescripcionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionIII.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionII.Text = ""
    End If
End Sub

Private Sub DescripcionIII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIII.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
        
            Sql1 = "Select *"
            Sql2 = " FROM EquipoFabrica"
            Sql3 = " Where EquipoFabrica.Codigo = " + "'" + Codigo.Text + "'"
            spEquipoFabrica = Sql1 + Sql2 + Sql3
            Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipoFabrica.RecordCount > 0 Then
                rstEquipoFabrica.Close
                Call Imprime_Datos
                    Else
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
        End If
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Codigo.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Anterior_Click()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo < " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveLast
            Codigo.Text = rstEquipoFabrica!Codigo
        End With
        rstEquipoFabrica.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos, Control y Instrucciones de Seguridad")
    End If
    
End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM EquipoFabrica"
    Sql3 = " Where EquipoFabrica.Codigo > " + "'" + Codigo.Text + "'"
    spEquipoFabrica = Sql1 + Sql2 + Sql3
    Set rstEquipoFabrica = db.OpenRecordset(spEquipoFabrica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipoFabrica.RecordCount > 0 Then
        With rstEquipoFabrica
            .MoveFirst
            Codigo.Text = rstEquipoFabrica!Codigo
        End With
        rstEquipoFabrica.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos, Control y Instrucciones de Seguridad")
    End If

End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""
    
End Sub

