VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgEquipos 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Equipos Usados"
   ClientHeight    =   6900
   ClientLeft      =   300
   ClientTop       =   1005
   ClientWidth     =   11430
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   11430
   Begin VB.TextBox PoeLimpieza 
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
      Left            =   6840
      MaxLength       =   20
      TabIndex        =   23
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Identificacion 
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
      TabIndex        =   21
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox Poe 
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
      TabIndex        =   19
      Top             =   1320
      Width           =   2535
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
      Top             =   960
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
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6000
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
      TabIndex        =   16
      Top             =   6000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   1920
      TabIndex        =   5
      Top             =   3240
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
         MouseIcon       =   "equipos.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":030A
         ToolTipText     =   "Confirma la Impresion"
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Cancela 
         Height          =   480
         Left            =   4320
         MouseIcon       =   "equipos.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":0A56
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
      Top             =   2160
      Width           =   3015
      Begin VB.Image Anterior 
         Height          =   480
         Left            =   840
         MouseIcon       =   "equipos.frx":0E98
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":11A2
         ToolTipText     =   "Registro Anterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Siguiente 
         Height          =   480
         Left            =   1560
         MouseIcon       =   "equipos.frx":15E4
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":18EE
         ToolTipText     =   "Registro Posterior"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Ultimo 
         Height          =   480
         Left            =   2280
         MouseIcon       =   "equipos.frx":1D30
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":203A
         ToolTipText     =   "Ultimo Registro"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Primer 
         Height          =   480
         Left            =   240
         MouseIcon       =   "equipos.frx":247C
         MousePointer    =   99  'Custom
         Picture         =   "equipos.frx":2786
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
      Top             =   3240
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
      ReportFileName  =   "equipo.rpt"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   3975
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   3015
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5318
      _Version        =   393216
      BackColor       =   16777152
   End
   Begin VB.Label lblLabels 
      Caption         =   "Poe Limpieza"
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
      Index           =   4
      Left            =   4800
      TabIndex        =   24
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Identificacion"
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
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      Caption         =   "Poe Operacion"
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
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Image Lista 
      Height          =   480
      Left            =   3720
      MouseIcon       =   "equipos.frx":2BC8
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":2ED2
      ToolTipText     =   "Impresion "
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   2040
      MouseIcon       =   "equipos.frx":3714
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":3A1E
      ToolTipText     =   "Limpia la pantalla"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   360
      MouseIcon       =   "equipos.frx":4260
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":456A
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   1200
      MouseIcon       =   "equipos.frx":4DAC
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":50B6
      ToolTipText     =   "Elimina el Registro"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   4560
      MouseIcon       =   "equipos.frx":58F8
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":5C02
      ToolTipText     =   "Salida"
      Top             =   2400
      Width           =   480
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   2880
      MouseIcon       =   "equipos.frx":6444
      MousePointer    =   99  'Custom
      Picture         =   "equipos.frx":674E
      ToolTipText     =   "Consulta de Datos"
      Top             =   2400
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
      Caption         =   "Codigo de Equipo"
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
Attribute VB_Name = "PrgEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstEquipo As Recordset
Dim spEquipo As String
Dim ZClave(10000, 2) As String

Sub Verifica_datos()
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "0"
    End If
End Sub

Sub Imprime_Datos()
    Sql1 = "Select *"
    Sql2 = " FROM Equipo"
    Sql3 = " Where Equipo.Codigo = " + "'" + Codigo.Text + "'"
    spEquipo = Sql1 + Sql2 + Sql3
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        Descripcion.Text = Trim(rstEquipo!Descripcion)
        DescripcionII.Text = Trim(rstEquipo!DescripcionII)
        Poe.Text = IIf(IsNull(rstEquipo!Poe), "", rstEquipo!Poe)
        Poe.Text = Trim(Poe.Text)
        PoeLimpieza.Text = IIf(IsNull(rstEquipo!PoeLimpieza), "", rstEquipo!PoeLimpieza)
        PoeLimpieza.Text = Trim(PoeLimpieza.Text)
        Identificacion.Text = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
        Identificacion.Text = Trim(Identificacion.Text)
        rstEquipo.Close
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
    
    Listado.SQLQuery = "SELECT Equipo.Codigo, Equipo.Descripcion, Equipo.DescripcionII, Equipo.Poe, Equipo.Identificacion, Equipo.PoeLimpieza " _
                + "From " _
                + DSQ + ".dbo.Equipo Equipo " _
                + "Where " _
                + "Equipo.Codigo >= " + Desde.Text + " AND " _
                + "Equipo.Codigo <= " + Hasta.Text
    
    Listado.Connect = Connect()
    
    Listado.GroupSelectionFormula = "{Equipo.Codigo} in " + Desde.Text + " to " + Hasta.Text
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
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + Codigo.Text + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
        
            ZZDescripcion = Trim(rstEquipo!Descripcion)
            ZZDescripcionII = Trim(rstEquipo!DescripcionII)
            ZZPoe = IIf(IsNull(rstEquipo!Poe), "", rstEquipo!Poe)
            ZZPoe = Trim(ZZPoe)
            ZZPoeLimpieza = IIf(IsNull(rstEquipo!PoeLimpieza), "", rstEquipo!PoeLimpieza)
            ZZPoeLimpieza = Trim(ZZPoeLimpieza)
            ZZIdentificacion = IIf(IsNull(rstEquipo!Identificacion), "", rstEquipo!Identificacion)
            ZZIdentificacion = Trim(ZZIdentificacion)
    
    
            XDescripcionI = "Utilizar Equipo : " + ZZIdentificacion
            XXDescripcionI = "Operar Equipo segun POE " + ZZPoe
            XXXDescripcionI = "Se debe realizar la limpieza del equipo segun POE : " + Trim(ZZPoeLimpieza)
    
            XDescripcionII = "Utilizar Equipo : " + Identificacion.Text
            XXDescripcionII = "Operar Equipo segun POE " + Poe.Text
            XXXDescripcionII = "Se debe realizar la limpieza del equipo segun POE : " + Trim(PoeLimpieza.Text)
            
            rstEquipo.Close
            
            If Trim(XDescripcionI) <> Trim(XDescripcionII) Then
                ZSql = ""
                ZSql = ZSql + "UPDATE CargaIII SET "
                ZSql = ZSql + " Descripcion = " + "'" + XDescripcionII + "'"
                ZSql = ZSql + " Where Descripcion = " + "'" + XDescripcionI + "'"
                ZSql = ZSql + " and Equipo = " + "'" + Codigo.Text + "'"
                spCargaIII = ZSql
                Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            
            If Trim(XXDescripcionI) <> Trim(XXDescripcionII) Then
                ZSql = ""
                ZSql = ZSql + "UPDATE CargaIII SET "
                ZSql = ZSql + " Descripcion = " + "'" + XXDescripcionII + "'"
                ZSql = ZSql + " Where Descripcion = " + "'" + XXDescripcionI + "'"
                ZSql = ZSql + " and Equipo = " + "'" + Codigo.Text + "'"
                spCargaIII = ZSql
                Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            If Trim(XXXDescripcionI) <> Trim(XXXDescripcionII) Then
                
                Erase ZClave
                ZLugar = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM CargaIII"
                ZSql = ZSql + " Where CargaIII.Descripcion LIKE " + "'" + "%" + XXXDescripcionI + "%" + "'"
                ZSql = ZSql + " Order by Clave"
                spCargaIII = ZSql
                Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaIII.RecordCount > 0 Then
                    With rstCargaIII
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                
                                If rstCargaIII!Equipo = Val(Codigo.Text) Then
                                
                                    ZLugar = ZLugar + 1
                                    ZClave(ZLugar, 1) = rstCargaIII!Clave
                                    ZClave(ZLugar, 2) = rstCargaIII!Metodo
                                    
                                End If
                            
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstCargaIII.Close
                End If
                
                
                
                For Ciclo = 1 To ZLugar
                
                    ZZZDescripcion = XXXDescripcionII + " Metodo:" + ZClave(Ciclo, 2)
                    ZZZDescripcion = Left$(ZZZDescripcion, 70)
                    
                    ZSql = ""
                    ZSql = ZSql + "UPDATE CargaIII SET "
                    ZSql = ZSql + " Descripcion = " + "'" + ZZZDescripcion + "'"
                    ZSql = ZSql + " Where Clave = " + "'" + ZClave(Ciclo, 1) + "'"
                    spCargaIII = ZSql
                    Set rstCargaIII = db.OpenRecordset(spCargaIII, dbOpenSnapshot, dbSQLPassThrough)
                
                Next Ciclo
                
            End If
    
        End If
    
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + Codigo.Text + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            rstEquipo.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE Equipo SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "',"
            ZSql = ZSql + " DescripcionII = " + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + " Poe = " + "'" + Poe.Text + "',"
            ZSql = ZSql + " PoeLimpieza = " + "'" + PoeLimpieza.Text + "',"
            ZSql = ZSql + " Identificacion = " + "'" + Identificacion.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spEquipo = ZSql
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO Equipo ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "DescripcionII ,"
            ZSql = ZSql + "Poe ,"
            ZSql = ZSql + "PoeLimpieza ,"
            ZSql = ZSql + "Identificacion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "',"
            ZSql = ZSql + "'" + DescripcionII.Text + "',"
            ZSql = ZSql + "'" + Poe.Text + "',"
            ZSql = ZSql + "'" + PoeLimpieza.Text + "',"
            ZSql = ZSql + "'" + Identificacion.Text + "')"
            spEquipo = ZSql
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        End If
    
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Val(Codigo.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM Equipo"
        Sql3 = " Where Equipo.Codigo = " + "'" + Codigo.Text + "'"
        spEquipo = Sql1 + Sql2 + Sql3
        Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEquipo.RecordCount > 0 Then
            rstEquipo.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE Equipo"
                Sql2 = " Where Codigo = " + "'" + Codigo.Text + "'"
                spEquipo = Sql1 + Sql2
                Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
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
    Poe.Text = ""
    PoeLimpieza.Text = ""
    Identificacion.Text = ""

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Equipo"
    spEquipo = Sql1 + Sql2
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        rstEquipo.MoveLast
        ZCodigo = IIf(IsNull(rstEquipo!CodigoMayor), "0", rstEquipo!CodigoMayor)
        Codigo.Text = ZCodigo + 1
        rstEquipo.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
    
    Codigo.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    Call CmdLimpiar_Click
    PrgEquipos.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Anterior_Click()
    Sql1 = "Select *"
    Sql2 = " FROM Equipo"
    Sql3 = " Where Equipo.Codigo < " + "'" + Codigo.Text + "'"
    spEquipo = Sql1 + Sql2 + Sql3
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        With rstEquipo
            .MoveLast
            Codigo.Text = rstEquipo!Codigo
        End With
        rstEquipo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos Usados")
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
        DescripcionII.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Private Sub DescripcionII_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Poe.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionII.Text = ""
    End If
End Sub

Private Sub Poe_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        PoeLimpieza.SetFocus
    End If
    If KeyAscii = 27 Then
        Poe.Text = ""
    End If
End Sub

Private Sub PoeLimpieza_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Identificacion.SetFocus
    End If
    If KeyAscii = 27 Then
        PoeLimpieza.Text = ""
    End If
End Sub

Private Sub Identificacion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion.SetFocus
    End If
    If KeyAscii = 27 Then
        Identificacion.Text = ""
    End If
End Sub

Private Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Equipo"
            Sql3 = " Where Equipo.Codigo = " + "'" + Codigo.Text + "'"
            spEquipo = Sql1 + Sql2 + Sql3
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                rstEquipo.Close
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
     Ayuda.Visible = False
     Opcion.Clear

     Opcion.AddItem "Equipos"

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
            Sql2 = " FROM Equipo"
            Sql3 = " Order by Equipo.Codigo"
            spEquipo = Sql1 + Sql2 + Sql3
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                With rstEquipo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstEquipo!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstEquipo!Descripcion
                            IngresaItem = rstEquipo!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEquipo.Close
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
    Sql2 = " FROM Equipo"
    spEquipo = Sql1 + Sql2
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        rstEquipo.MoveFirst
        Codigo.Text = rstEquipo!CodigoMenor
        rstEquipo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Ultimo_Click()

    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Equipo"
    spEquipo = Sql1 + Sql2
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        rstEquipo.MoveLast
        Codigo.Text = rstEquipo!CodigoMayor
        rstEquipo.Close
        Call Imprime_Datos
        Codigo.SetFocus
    End If
    
 End Sub

Private Sub Siguiente_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Equipo"
    Sql3 = " Where Equipo.Codigo > " + "'" + Codigo.Text + "'"
    spEquipo = Sql1 + Sql2 + Sql3
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        With rstEquipo
            .MoveFirst
            Codigo.Text = rstEquipo!Codigo
        End With
        rstEquipo.Close
        Call Imprime_Datos
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Equipos Usados")
    End If

End Sub

Sub Form_Load()

    Codigo.Text = ""
    Descripcion.Text = ""
    DescripcionII.Text = ""
    Poe.Text = ""
    PoeLimpieza.Text = ""
    Identificacion.Text = ""
    
    Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Sql2 = " FROM Equipo"
    spEquipo = Sql1 + Sql2
    Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEquipo.RecordCount > 0 Then
        rstEquipo.MoveLast
        ZCodigo = IIf(IsNull(rstEquipo!CodigoMayor), "0", rstEquipo!CodigoMayor)
        Codigo.Text = ZCodigo + 1
        rstEquipo.Close
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
            Sql1 = "Select *"
            Sql2 = " FROM Equipo"
            Sql3 = " Order by Equipo.Codigo"
            spEquipo = Sql1 + Sql2 + Sql3
            Set rstEquipo = db.OpenRecordset(spEquipo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEquipo.RecordCount > 0 Then
                With rstEquipo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstEquipo!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstEquipo!Descripcion, aa, WEspacios) Then
                                    LugarAyuda = LugarAyuda + 1
                                    Pantalla.Row = LugarAyuda
                                    Pantalla.Col = 1
                                    Pantalla.Text = rstEquipo!Codigo
                                    Pantalla.Col = 2
                                    Pantalla.Text = rstEquipo!Descripcion
                                    IngresaItem = rstEquipo!Codigo
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
                rstEquipo.Close
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
    Opcion.AddItem "Equipos"
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
                        Pantalla.Text = "Equipo"
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





