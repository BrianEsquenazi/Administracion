VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgFraseH 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Frases H"
   ClientHeight    =   7950
   ClientLeft      =   300
   ClientTop       =   1005
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   7950
   ScaleWidth      =   11430
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   975
      Left            =   8640
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Observa 
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
      TabIndex        =   20
      Top             =   2520
      Width           =   4095
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
      TabIndex        =   19
      Top             =   1560
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
      TabIndex        =   18
      Top             =   1080
      Width           =   9015
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
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6720
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
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   2040
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox Hasta 
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
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   11
         Text            =   " "
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Desde 
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
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   10
         Text            =   " "
         Top             =   360
         Width           =   2055
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
         MouseIcon       =   "fraseh.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "fraseh.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":0A56
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
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   1215
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
         Left            =   360
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5280
      TabIndex        =   14
      Top             =   2880
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "fraseh.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "fraseh.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "fraseh.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "fraseh.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "fraseh.frx":2786
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
      Top             =   3960
      Visible         =   0   'False
      Width           =   8175
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
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   2175
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4800
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "MetodoFiltrado.rpt"
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
      MaxLength       =   100
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
      Top             =   4320
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   4320
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   327680
      BackColor       =   16777152
   End
   Begin VB.Label Obserfv 
      Caption         =   "Observacion"
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
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "fraseh.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "fraseh.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "fraseh.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "fraseh.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "fraseh.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "fraseh.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "fraseh.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   3120
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
Attribute VB_Name = "PrgFraseH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstFraseH As Recordset
Dim spFraseH As String
Dim Wactual As String


Sub Verifica_datos()
End Sub

Sub Imprime_Datos()
    Sql1 = "Select *"
    Sql2 = " FROM FraseH"
    Sql3 = " Where FraseH.Codigo = " + "'" + Codigo.Text + "'"
    spFraseH = Sql1 + Sql2 + Sql3
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        Descripcion.Text = Trim(rstFraseH!Descripcion)
        DescripcionII.Text = Trim(rstFraseH!DescripcionII)
        DescripcionIII.Text = Trim(rstFraseH!DescripcionIII)
        Observa.Text = Trim(rstFraseH!Observa)
        rstFraseH.Close
    End If
End Sub

Private Sub Acepta_Click()
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT FraseH.Codigo, FraseH.Descripcion " _
                + "From " _
                + DSQ + ".dbo.FraseH FraseH " _
                + "Where " _
                + "FraseH.Codigo >= " + Desde.Text + " AND " _
                + "FraseH.Codigo <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{FraseH.Codigo} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
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
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM FraseH"
        ZSql = ZSql + " Where FraseH.Codigo = " + "'" + Codigo.Text + "'"
        spFraseH = ZSql
        Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        If rstFraseH.RecordCount > 0 Then
            rstFraseH.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE FraseH SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + " DescripcionIII = " + "'" + DescripcionIII.Text + "',"
            ZSql = ZSql + " Observa = " + "'" + Observa.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spFraseH = ZSql
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO FraseH ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "DescripcionIII ,"
            ZSql = ZSql + "Observa )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + "'" + Observa.Text + "')"
            spFraseH = ZSql
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
    
End Sub

Private Sub cmdDelete_Click()

    If Codigo.Text <> "" Then
        Sql1 = "Select *"
        Sql2 = " FROM FraseH"
        Sql3 = " Where FraseH.Codigo = " + "'" + Codigo.Text + "'"
        spFraseH = Sql1 + Sql2 + Sql3
        Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        If rstFraseH.RecordCount > 0 Then
            rstFraseH.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE FraseH"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spFraseH = Sql1 + Sql2
                Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
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
    Observa.Text = ""

    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
      
    
    PrgFraseH.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    Sql1 = "Select *"
    Sql2 = " FROM FraseH"
    Sql3 = " Where FraseH.Codigo < " + "'" + Codigo.Text + "'"
    spFraseH = Sql1 + Sql2 + Sql3
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        With rstFraseH
            .MoveLast
            Codigo.Text = rstFraseH!Codigo
        End With
        rstFraseH.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        a% = MsgBox(m$, 0, "Archivo de Frases H")
    End If
End Sub

Private Sub Command1_Click()

    Dim ZZVector(1000, 2) As String

    Erase ZZVector
    ZZLugar = 0

    Sql1 = "Select *"
    Sql2 = " FROM FraseH"
    Sql3 = " Order by FraseH.Codigo"
    spFraseH = Sql1 + Sql2 + Sql3
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        With rstFraseH
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZZLugar = ZZLugar + 1
                    ZZVector(ZZLugar, 1) = rstFraseH!Codigo
                    ZZVector(ZZLugar, 2) = rstFraseH!Descripcion
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFraseH.Close
    End If
    
    For Ciclo = 1 To ZZLugar
    
        ZZCodigo = Trim(ZZVector(Ciclo, 1))
        ZZDescripcion = Trim(ZZVector(Ciclo, 2))
        ZZDescripcionII = ""
        ZZDescripcionIII = ""
        
        If Len(ZZDescripcion) <= 100 Then
            
                Else
                
            For Cicla = 100 To 1 Step -1
                If Mid(ZZDescripcion, Cicla, 1) = Space(1) Then
                    ZZDescripcionII = Mid(ZZDescripcion, Cicla + 1, 250)
                    ZZDescripcion = Mid(ZZDescripcion, 1, Cicla)
                    Exit For
                End If
            Next Cicla
            
            If Len(ZZDescripcionII) > 100 Then
            
                For Cicla = 100 To 1 Step -1
                    If Mid(ZZDescripcionII, Cicla, 1) = Space(1) Then
                        ZZDescripcionIII = Mid(ZZDescripcionII, Cicla + 1, 250)
                        ZZDescripcionII = Mid(ZZDescripcionII, 1, Cicla)
                        Exit For
                    End If
                Next Cicla
                
            End If
            
        End If
        
        aa = Len(ZZDescripcion)
        aaII = Len(ZZDescripcionII)
        aaIII = Len(ZZDescripcionIII)
        
        ZSql = ""
        ZSql = ZSql + "UPDATE FraseH SET "
        ZSql = ZSql + " Descripcion = " + "'" + ZZDescripcion + "',"
        ZSql = ZSql + " DescripcionII = " + "'" + ZZDescripcionII + "',"
        ZSql = ZSql + " DescripcionIII = " + "'" + ZZDescripcionIII + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + ZZCodigo + "'"
        spFraseH = ZSql
        Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo


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
        Observa.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIII.Text = ""
    End If
End Sub

Private Sub Observa_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Observa.Text = ""
    End If
End Sub







Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Sql1 = "Select *"
            Sql2 = " FROM FraseH"
            Sql3 = " Where FraseH.Codigo = " + "'" + Codigo.Text + "'"
            spFraseH = Sql1 + Sql2 + Sql3
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseH.RecordCount > 0 Then
                rstFraseH.Close
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
End Sub

Private Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    If KeyAscii = 27 Then
        Desde.Text = ""
    End If
End Sub

Private Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    If KeyAscii = 27 Then
        Hasta.Text = ""
    End If
End Sub

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Frases H"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    On Error GoTo WError
    
    Opcion.Visible = False
     
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM FraseH"
            Sql3 = " Order by FraseH.Codigo"
            spFraseH = Sql1 + Sql2 + Sql3
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseH.RecordCount > 0 Then
                With rstFraseH
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstFraseH!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstFraseH!Descripcion
                            IngresaItem = rstFraseH!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFraseH.Close
            End If
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    Sql1 = "Select Min(Codigo) as [CodigoMenor]"
    Sql2 = " FROM FraseH"
    spFraseH = Sql1 + Sql2
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        rstFraseH.MoveFirst
        Codigo.Text = rstFraseH!CodigoMenor
        rstFraseH.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM FraseH"
    spFraseH = Sql1 + Sql2
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        rstFraseH.MoveLast
        Codigo.Text = rstFraseH!CodigoMayor
        rstFraseH.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM FraseH"
    Sql3 = " Where FraseH.Codigo > " + "'" + Codigo.Text + "'"
    spFraseH = Sql1 + Sql2 + Sql3
    Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
    If rstFraseH.RecordCount > 0 Then
        With rstFraseH
            .MoveFirst
            Codigo.Text = rstFraseH!Codigo
        End With
        rstFraseH.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        a% = MsgBox(m$, 0, "Archivo de Frases H")
    End If

End Sub

Sub Form_Load()

    
    
    
    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    DescripcionIII.Text = ""
    Observa.Text = ""
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM FraseH"
            Sql3 = " Order by FraseH.Codigo"
            spFraseH = Sql1 + Sql2 + Sql3
            Set rstFraseH = db.OpenRecordset(spFraseH, dbOpenSnapshot, dbSQLPassThrough)
            If rstFraseH.RecordCount > 0 Then
                With rstFraseH
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstFraseH!Descripcion) - WEspacios
                            For aa = 1 To Da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstFraseH!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstFraseH!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstFraseH!Descripcion
                                    IngresaItem = rstFraseH!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstFraseH.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Codigo_DblClick()

    Opcion.Clear
    Opcion.AddItem "Frases H"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Limpia_Ayuda()

    Pantalla.Clear
    Pantalla.Font.Bold = True
    
    ' Establesco loa Valores de la pantalla
    
    XIndice = Opcion.ListIndex
    Select Case XIndice
        Case 0
            Pantalla.FixedCols = 1
            Pantalla.Cols = 3
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
    End Select
    
    Pantalla.ColWidth(0) = 200
    Pantalla.Row = 0
    
    Select Case XIndice
        Case 0
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
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    
    Pantalla.Row = 0
    For Ciclo = 1 To Pantalla.Cols - 1
        Pantalla.Col = Ciclo
        WTitulo(Ciclo).Text = Pantalla.Text
        WTitulo(Ciclo).Left = Pantalla.CellLeft + Pantalla.Left
        WTitulo(Ciclo).Top = Pantalla.CellTop + Pantalla.Top
        WTitulo(Ciclo).Width = Pantalla.CellWidth
        WTitulo(Ciclo).Height = Pantalla.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA pantalla
    
    WAncho = 400
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub





