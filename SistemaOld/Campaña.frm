VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCampa�a 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Campa�as Publicitarias"
   ClientHeight    =   6540
   ClientLeft      =   1560
   ClientTop       =   1095
   ClientWidth     =   8685
   LinkTopic       =   "Form2"
   ScaleHeight     =   6540
   ScaleWidth      =   8685
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
      Index           =   3
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   5520
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
      Index           =   2
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   5520
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
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   12
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
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   11
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
         Left            =   1800
         TabIndex        =   10
         Top             =   1320
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
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Image Acepta 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "Campa�a.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "Campa�a.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":0A56
         ToolTipText     =   "Cancela la Impresion"
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta Forma de Contacto"
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
         Left            =   240
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Forma de Contacto"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   5280
      TabIndex        =   17
      Top             =   1560
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "Campa�a.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "Campa�a.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "Campa�a.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "Campa�a.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "Campa�a.frx":2786
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
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Contacto 
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
      TabIndex        =   2
      Text            =   " "
      Top             =   1080
      Width           =   855
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
      ReportFileName  =   "WCampa�a.rpt"
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
      Left            =   5760
      TabIndex        =   5
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
      Width           =   4935
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
      TabIndex        =   15
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   3000
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
      MouseIcon       =   "Campa�a.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "Campa�a.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "Campa�a.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "Campa�a.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "Campa�a.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "Campa�a.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "Campa�a.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label DesContacto 
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
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Forma de Contacto"
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
      TabIndex        =   13
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nombre"
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
      TabIndex        =   4
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Campa�a"
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
      TabIndex        =   3
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgCampa�a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCampa�a As Recordset
Dim spCampa�a As String
Dim rstContacto As Recordset
Dim spContacto As String

Sub Imprime_Descripcion()
    Sql1 = "Select *"
    Sql2 = " FROM Contacto"
    Sql3 = " Where Contacto.Codigo = " + "'" + Contacto.Text + "'"
    spContacto = Sql1 + Sql2 + Sql3
    Set rstContacto = db.OpenRecordset(spContacto, dbOpenSnapshot, dbSQLPassThrough)
    If rstContacto.RecordCount > 0 Then
        DesContacto.Caption = rstContacto!Descripcion
        rstContacto.Close
            Else
        DesContacto.Caption = ""
    End If
End Sub

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
    If Val(Contacto.Text) = 0 Then
        Contacto.Text = "0"
    End If
End Sub

Sub Format_datos()
    Rem Habitantes.Text = Pusing("###,###,###", Habitantes.Text)
End Sub

Sub Imprime_Datos()
    Sql1 = "Select *"
    Sql2 = " FROM Campa�a"
    Sql3 = " Where Campa�a.Codigo = " + "'" + Codigo.Text + "'"
    spCampa�a = Sql1 + Sql2 + Sql3
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        Descripcion.Text = rstCampa�a!Descripcion
        Contacto.Text = rstCampa�a!Contacto
        rstCampa�a.Close
        Call Format_datos
        Call Imprime_Descripcion
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
    
    Listado.SQLQuery = "SELECT Campa�a.Codigo, Campa�a.Descripcion, Campa�a.Contacto, " _
                    + "Contacto.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.Campa�a Campa�a, " _
                    + DSQ + ".dbo.Contacto Contacto " _
                    + "Where " _
                    + "Campa�a.Contacto = Contacto.Codigo AND " _
                    + "Campa�a.Contacto >= " + Desde.Text + " AND " _
                    + "Campa�a.Contacto <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Campa�a.Contacto} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    Codigo.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub

Private Sub CANCELA_Click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Val(Codigo.Text) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Campa�a"
        Sql3 = " Where Campa�a.Codigo = " + "'" + Codigo.Text + "'"
        spCampa�a = Sql1 + Sql2 + Sql3
        Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
        If rstCampa�a.RecordCount > 0 Then
            rstCampa�a.Close
            Sql1 = "UPDATE Campa�a SET "
            Sql2 = " Descripcion = " + "'" + Descripcion.Text + "',"
            Sql3 = " Contacto = " + "'" + Contacto.Text + "'"
            Sql4 = " Where Codigo = " + "'" + Codigo.Text + "'"
            spCampa�a = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
                Else
            Sql1 = "INSERT INTO Campa�a ("
            Sql2 = "Codigo ,"
            Sql3 = "Descripcion ,"
            Sql4 = "Contacto )"
            Sql5 = "Values ("
            Sql6 = "'" + Codigo.Text + "',"
            Sql7 = "'" + Descripcion.Text + "',"
            Sql8 = "'" + Contacto.Text + "')"
            spCampa�a = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        If WTipoGrabacion = 0 Then
    
            txtOdbc = "Call02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
            Sql1 = "Select *"
            Sql2 = " FROM Campa�a"
            Sql3 = " Where Campa�a.Codigo = " + "'" + Codigo.Text + "'"
            spCampa�a = Sql1 + Sql2 + Sql3
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
            If rstCampa�a.RecordCount > 0 Then
                rstCampa�a.Close
                Sql1 = "UPDATE Campa�a SET "
                Sql2 = " Descripcion = " + "'" + Descripcion.Text + "',"
                Sql3 = " Contacto = " + "'" + Contacto.Text + "'"
                Sql4 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spCampa�a = Sql1 + Sql2 + Sql3 + Sql4
                Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                Sql1 = "INSERT INTO Campa�a ("
                Sql2 = "Codigo ,"
                Sql3 = "Descripcion ,"
                Sql4 = "Contacto )"
                Sql5 = "Values ("
                Sql6 = "'" + Codigo.Text + "',"
                Sql7 = "'" + Descripcion.Text + "',"
                Sql8 = "'" + Contacto.Text + "')"
                spCampa�a = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8
                Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            txtOdbc = "Call01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Val(Codigo.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM Campa�a"
        Sql3 = " Where Campa�a.Codigo = " + "'" + Codigo.Text + "'"
        spCampa�a = Sql1 + Sql2 + Sql3
        Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
        If rstCampa�a.RecordCount > 0 Then
            rstCampa�a.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                Sql1 = "DELETE Campa�a"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spCampa�a = Sql1 + Sql2
                Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
                
                If WTipoGrabacion = 0 Then
    
                    txtOdbc = "Call02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
                    Sql1 = "DELETE Campa�a"
                    Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                    spCampa�a = Sql1 + Sql2
                    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
        
                    txtOdbc = "Call01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                End If
                
                Call CmdLimpiar_Click
                
            End If
        End If
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Codigo.Text = ""
    Descripcion.Text = ""
    Contacto.Text = ""
    DesContacto.Caption = ""

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Campa�a"
    spCampa�a = Sql1 + Sql2
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        rstCampa�a.MoveLast
        Codigo.Text = rstCampa�a!CodigoMayor + 1
        rstCampa�a.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgCampa�a.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    Sql1 = "Select *"
    Sql2 = " FROM Campa�a"
    Sql3 = " Where Campa�a.Codigo < " + "'" + Codigo.Text + "'"
    spCampa�a = Sql1 + Sql2 + Sql3
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        With rstCampa�a
            .MoveLast
            Codigo.Text = rstCampa�a!Codigo
        End With
        rstCampa�a.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Campa�a")
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
        Contacto.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub Contacto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Contacto"
        Sql3 = " Where Contacto.Codigo = " + "'" + Contacto.Text + "'"
        spContacto = Sql1 + Sql2 + Sql3
        Set rstContacto = db.OpenRecordset(spContacto, dbOpenSnapshot, dbSQLPassThrough)
        If rstContacto.RecordCount > 0 Then
            DesContacto.Caption = rstContacto!Descripcion
            Descripcion.SetFocus
            rstContacto.Close
                Else
            Contacto.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Contacto.Text = ""
        DesContacto.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Campa�a"
            Sql3 = " Where Campa�a.Codigo = " + "'" + Codigo.Text + "'"
            spCampa�a = Sql1 + Sql2 + Sql3
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
            If rstCampa�a.RecordCount > 0 Then
                rstCampa�a.Close
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

Private Sub Consulta_Click()

     Pantalla.Visible = False
     WTitulo(1).Visible = False
     WTitulo(2).Visible = False
     WTitulo(3).Visible = False
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Campa�as"
     Opcion.AddItem "Forma de Contacto"

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
            Sql1 = "Select Campa�a.Codigo, Campa�a.Descripcion, Campa�a.Contacto, Contacto.Descripcion as [WDesContacto]"
            Sql2 = " FROM Campa�a, Contacto"
            Sql3 = " Where Campa�a.Contacto = Contacto.Codigo"
            Sql4 = " Order by Campa�a.Codigo"
            spCampa�a = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
            If rstCampa�a.RecordCount > 0 Then
                With rstCampa�a
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstCampa�a!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstCampa�a!Descripcion
                            Pantalla.Col = 3
                            Pantalla.Text = rstCampa�a!WDesContacto
                            IngresaItem = rstCampa�a!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCampa�a.Close
            End If
            
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Contacto"
            Sql3 = " Order BY Contacto"
            spContacto = Sql1 + Sql2
            Set rstContacto = db.OpenRecordset(spContacto, dbOpenSnapshot, dbSQLPassThrough)
            If rstContacto.RecordCount > 0 Then
                With rstContacto
                    .Index = "Codigo"
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = !Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = !Descripcion
                            IngresaItem = !Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstContacto.Close
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

Private Sub Pantalla_Click()

    Pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    
    Select Case XIndice
        Case 0
            Indice = Pantalla.Row - 1
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case 1
            Indice = Pantalla.Row - 1
            Contacto.Text = WIndice.List(Indice)
            Call Contacto_KeyPress(13)
            
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    Sql1 = "Select Min(Codigo) as [CodigoMenor]"
    Sql2 = " FROM Campa�a"
    spCampa�a = Sql1 + Sql2
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        rstCampa�a.MoveFirst
        Codigo.Text = rstCampa�a!CodigoMenor
        rstCampa�a.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Campa�a"
    spCampa�a = Sql1 + Sql2
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        rstCampa�a.MoveLast
        Codigo.Text = rstCampa�a!CodigoMayor
        rstCampa�a.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Campa�a"
    Sql3 = " Where Campa�a.Codigo > " + "'" + Codigo.Text + "'"
    spCampa�a = Sql1 + Sql2 + Sql3
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        With rstCampa�a
            .MoveFirst
            Codigo.Text = rstCampa�a!Codigo
        End With
        rstCampa�a.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Campa�a")
    End If

End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    Contacto.Text = ""
    DesContacto.Caption = ""
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Campa�a"
    spCampa�a = Sql1 + Sql2
    Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
    If rstCampa�a.RecordCount > 0 Then
        rstCampa�a.MoveLast
        Codigo.Text = rstCampa�a!CodigoMayor + 1
        rstCampa�a.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
    
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
            Sql1 = "Select Campa�a.Codigo, Campa�a.Descripcion, Campa�a.Contacto, Contacto.Descripcion as [WDesContacto]"
            Sql2 = " FROM Campa�a, Contacto"
            Sql3 = " Where Campa�a.Contacto = Contacto.Codigo"
            Sql4 = " Order by Campa�a.Codigo"
            spCampa�a = Sql1 + Sql2 + Sql3 + Sql4
            Set rstCampa�a = db.OpenRecordset(spCampa�a, dbOpenSnapshot, dbSQLPassThrough)
            If rstCampa�a.RecordCount > 0 Then
                With rstCampa�a
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCampa�a!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstCampa�a!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstCampa�a!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstCampa�a!Descripcion
                                    Pantalla.Col = 3
                                    Pantalla.Text = rstCampa�a!WDesContacto
                                    IngresaItem = rstCampa�a!Codigo
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
                rstCampa�a.Close
            End If
                
                
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM Contacto"
            Sql3 = " Order BY Contacto"
            spContacto = Sql1 + Sql2
            Set rstContacto = db.OpenRecordset(spContacto, dbOpenSnapshot, dbSQLPassThrough)
            If rstContacto.RecordCount > 0 Then
                With rstContacto
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstContacto!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstContacto!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstContacto!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstContacto!Descripcion
                                    IngresaItem = rstContacto!Codigo
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
                rstContacto.Close
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
    Opcion.AddItem "Campa�as"
    Opcion.AddItem "Forma de Contacto"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Contacto_DblClick()

    Opcion.Clear
    Opcion.AddItem "Campa�as"
    Opcion.AddItem "Forma de Contacto"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
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
            Pantalla.Cols = 4
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
        Case Else
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
                        Pantalla.Text = "Campa�a"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 3000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 3
                        Pantalla.Text = "Publicidad"
                        Pantalla.ColWidth(Ciclo) = 2000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                End Select
            Next Ciclo
            
        Case 1
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Tipo"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignRightCenter
                    Case 2
                        Pantalla.Text = "Nombre"
                        Pantalla.ColWidth(Ciclo) = 2500
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
               End Select
            Next Ciclo
        Case Else
            
    End Select
    
    Rem DESPILEGA LOS TITULOS
    
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    
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
    
    WAncho = 340
    For Ciclo = 0 To Pantalla.Cols - 1
        WAncho = WAncho + Pantalla.ColWidth(Ciclo)
    Next Ciclo
    Pantalla.Width = WAncho

    ' Size the columns.
    Font.Name = Pantalla.Font.Name
    Font.Size = Pantalla.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tama�o de las celdas
    Pantalla.AllowUserResizing = flexResizeBoth
    
    Pantalla.Col = 1
    Pantalla.Row = 1
    
End Sub





