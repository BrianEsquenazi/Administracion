VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgSectores 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Sectores"
   ClientHeight    =   5745
   ClientLeft      =   300
   ClientTop       =   1005
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   5745
   ScaleWidth      =   11430
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   4920
      Visible         =   0   'False
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
      Index           =   1
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
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
         MouseIcon       =   "SectorInve.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "SectorInve.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":0A56
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
      TabIndex        =   14
      Top             =   1080
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "SectorInve.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "SectorInve.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "SectorInve.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "SectorInve.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "SectorInve.frx":2786
         ToolTipText     =   "Primer Registro"
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.TextBox Ayuda 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Visible         =   0   'False
      Width           =   8175
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
      Top             =   120
      Width           =   855
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Sector.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Efluentes de Lavado"
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
      TabIndex        =   4
      Top             =   0
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
      MaxLength       =   50
      TabIndex        =   1
      Top             =   600
      Width           =   9015
   End
   Begin VB.ListBox Opcion 
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
      Left            =   1560
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "SectorInve.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "SectorInve.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "SectorInve.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "SectorInve.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "SectorInve.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   1320
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "SectorInve.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "SectorInve.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   1320
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
      Top             =   600
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
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgSectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstSectorInve As Recordset
Dim spSectorInve As String

Sub Verifica_datos()
    If Val(codigo.Text) = 0 Then
        codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()
    sql1 = "Select *"
    Sql2 = " FROM SectorInve"
    Sql3 = " Where SectorInve.Codigo = " + "'" + codigo.Text + "'"
    spSectorInve = sql1 + Sql2 + Sql3
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        descripcion.Text = Trim(rstSectorInve!descripcion)
        rstSectorInve.Close
    End If
End Sub

Private Sub Acepta_Click()
    If Val(Desde.Text) = 0 Then
         Desde.Text = "0"
    End If
    If Val(Hasta.Text) = 0 Then
         Hasta.Text = "0"
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT SectorInve.Codigo, SectorInve.Descripcion " _
                + "From " _
                + DSQ + ".dbo.SectorInve SectorInve " _
                + "Where " _
                + "SectorInve.Codigo >= " + Desde.Text + " AND " _
                + "SectorInve.Codigo <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{SectorInve.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(codigo.Text) <> 0 Then
    
        sql1 = "Select *"
        Sql2 = " FROM SectorInve"
        Sql3 = " Where SectorInve.Codigo = " + "'" + codigo.Text + "'"
        spSectorInve = sql1 + Sql2 + Sql3
        Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
        If rstSectorInve.RecordCount > 0 Then
            rstSectorInve.Close
            sql1 = "UPDATE SectorInve SET "
            Sql2 = " Descripcion = " + "'" + descripcion.Text + "'"
            Sql3 = " Where Codigo = " + "'" + codigo.Text + "'"
            spSectorInve = sql1 + Sql2 + Sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
                Else
            sql1 = "INSERT INTO SectorInve ("
            Sql2 = "Codigo ,"
            Sql3 = "Descripcion )"
            Sql4 = "Values ("
            Sql5 = "'" + codigo.Text + "',"
            Sql6 = "'" + descripcion.Text + "')"
            spSectorInve = sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        codigo.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Val(codigo.Text) <> 0 Then
        sql1 = "Select *"
        Sql2 = " FROM SectorInve"
        Sql3 = " Where SectorInve.Codigo = " + "'" + codigo.Text + "'"
        spSectorInve = sql1 + Sql2 + Sql3
        Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
        If rstSectorInve.RecordCount > 0 Then
            rstSectorInve.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                sql1 = "DELETE SectorInve"
                Sql2 = " Where Codigo = " + "'" + codigo.Text + "'"
                spSectorInve = sql1 + Sql2
                Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    
    codigo.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()

    codigo.Text = ""
    descripcion.Text = ""

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM SectorInve"
    spSectorInve = sql1 + Sql2
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        rstSectorInve.MoveLast
        ZCodigo = IIf(IsNull(rstSectorInve!CodigoMayor), "0", rstSectorInve!CodigoMayor)
        codigo.Text = ZCodigo + 1
        rstSectorInve.Close
    End If
    If Val(codigo.Text) = 0 Then
        codigo.Text = "1"
    End If
    
    codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgSectores.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    sql1 = "Select *"
    Sql2 = " FROM SectorInve"
    Sql3 = " Where SectorInve.Codigo < " + "'" + codigo.Text + "'"
    Sql4 = " Order by SectorInve.Codigo"
    spSectorInve = sql1 + Sql2 + Sql3 + Sql4
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        With rstSectorInve
            .MoveLast
            codigo.Text = rstSectorInve!codigo
        End With
        rstSectorInve.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de SectorInve de Lavado")
    End If
End Sub

Private Sub Lista_Click()
    Desde.Text = ""
    Hasta.Text = ""
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
        descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        descripcion.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(codigo.Text) <> 0 Then
            sql1 = "Select *"
            Sql2 = " FROM SectorInve"
            Sql3 = " Where SectorInve.Codigo = " + "'" + codigo.Text + "'"
            spSectorInve = sql1 + Sql2 + Sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                rstSectorInve.Close
                Call Imprime_Datos
                    Else
                WCodigo = codigo.Text
                CmdLimpiar_Click
                codigo.Text = WCodigo
            End If
        End If
        descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        codigo.Text = ""
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

Private Sub Consulta_Click()

     pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Sectores"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            Sql2 = " FROM SectorInve"
            Sql3 = " Order by SectorInve.Codigo"
            spSectorInve = sql1 + Sql2 + Sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                With rstSectorInve
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 1
                            pantalla.Text = rstSectorInve!codigo
                            pantalla.Col = 2
                            pantalla.Text = rstSectorInve!descripcion
                            IngresaItem = rstSectorInve!codigo
                            windice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSectorInve.Close
            End If
            
        Case Else
    End Select
            
    pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla_Click()

    pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = pantalla.Row - 1
            codigo.Text = windice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    sql1 = "Select Min(Codigo) as [CodigoMenor]"
    Sql2 = " FROM SectorInve"
    spSectorInve = sql1 + Sql2
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        rstSectorInve.MoveFirst
        codigo.Text = rstSectorInve!CodigoMenor
        rstSectorInve.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM SectorInve"
    spSectorInve = sql1 + Sql2
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        rstSectorInve.MoveLast
        codigo.Text = rstSectorInve!CodigoMayor
        rstSectorInve.Close
        Call Imprime_Datos
        codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    sql1 = "Select *"
    Sql2 = " FROM SectorInve"
    Sql3 = " Where SectorInve.Codigo > " + "'" + codigo.Text + "'"
    Sql4 = " Order by SectorInve.Codigo"
    spSectorInve = sql1 + Sql2 + Sql3 + Sql4
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        With rstSectorInve
            .MoveFirst
            codigo.Text = rstSectorInve!codigo
        End With
        rstSectorInve.Close
        Call Imprime_Datos
        codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de SectorInve de Lavado")
    End If

End Sub

Sub Form_Load()

    codigo.Text = ""
    descripcion.Text = ""
    
    sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM SectorInve"
    spSectorInve = sql1 + Sql2
    Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
    If rstSectorInve.RecordCount > 0 Then
        rstSectorInve.MoveLast
        ZCodigo = IIf(IsNull(rstSectorInve!CodigoMayor), "0", rstSectorInve!CodigoMayor)
        codigo.Text = ZCodigo + 1
        rstSectorInve.Close
    End If
    
    If Val(codigo.Text) = 0 Then
        codigo.Text = "1"
    End If
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            sql1 = "Select *"
            Sql2 = " FROM SectorInve"
            Sql3 = " Order by SectorInve.Codigo"
            spSectorInve = sql1 + Sql2 + Sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                With rstSectorInve
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstSectorInve!descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstSectorInve!descripcion, aa, WEspacios) Then
                                    Lugarayuda = Lugarayuda + 1
                                    pantalla.Row = Lugarayuda
                                    pantalla.Col = 1
                                    pantalla.Text = rstSectorInve!codigo
                                    pantalla.Col = 2
                                    pantalla.Text = rstSectorInve!descripcion
                                    IngresaItem = rstSectorInve!codigo
                                    windice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSectorInve.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Codigo_DblClick()

    Rem Opcion.Clear
    Rem Opcion.AddItem "SectorInve de Lavado"
    Rem Rem Opcion.Visible = True
    Rem Opcion.ListIndex = 0
    
    Rem Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    pantalla.Clear
    pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            pantalla.FixedCols = 1
            pantalla.Cols = 3
            pantalla.FixedRows = 1
            pantalla.Rows = 10001
    End Select
    
    pantalla.ColWidth(0) = 200
    pantalla.Row = 0
    
    Select Case XIndice
        Case 0
            For Ciclo = 1 To pantalla.Cols - 1
                pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        pantalla.Text = "SectorInve"
                        pantalla.ColWidth(Ciclo) = 1000
                        pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        pantalla.Text = "Nombre"
                        pantalla.ColWidth(Ciclo) = 6000
                        pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    pantalla.Row = 0
    For Ciclo = 1 To pantalla.Cols - 1
        pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = pantalla.Text
        WTitulo(Ciclo).Left = pantalla.CellLeft + pantalla.Left
        WTitulo(Ciclo).Top = pantalla.CellTop + pantalla.Top
        WTitulo(Ciclo).Width = pantalla.CellWidth
        WTitulo(Ciclo).Height = pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To pantalla.Cols - 1
        WAncho = WAncho + pantalla.ColWidth(Ciclo)
    Next Ciclo
    pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = pantalla.Font.Name
    Font.Size = pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    pantalla.AllowUserResizing = flexResizeBoth
    
    pantalla.Col = 1
    pantalla.Row = 1
    
End Sub





