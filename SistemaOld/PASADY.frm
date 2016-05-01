VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPasaDy 
   AutoRedraw      =   -1  'True
   Caption         =   "Pasa Dy"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   4815
      Begin MSMask.MaskEdBox ArticuloDy 
         Height          =   300
         Left            =   1680
         TabIndex        =   6
         Top             =   720
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Articulo 
         Height          =   300
         Left            =   1680
         TabIndex        =   0
         Top             =   360
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
         Mask            =   "AA-###-###"
         PromptChar      =   " "
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
         Left            =   3480
         TabIndex        =   5
         Top             =   360
         Width           =   975
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
         Left            =   3480
         TabIndex        =   4
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "ArticuloDy"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Articulo"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
End
Attribute VB_Name = "PrgPasaDy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WInicial As Double
Private WOrden As String
Private WClave As String
Private WDescripcion As String
Private WSaldo As Double

Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstComposicion As Recordset
Dim spComposicion As String
Dim rstCotiza As Recordset
Dim spCotiza As String
Dim rstEspecificaciones As Recordset
Dim spEspecificaciones As String
Dim rstMovguia As Recordset
Dim spMovguia As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstMovlab As Recordset
Dim spMovlab As String
Dim rstMovvar As Recordset
Dim spMovvar As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstPrestamo As Recordset
Dim spPrestamo As String
Dim rstPrueart As Recordset
Dim spPrueart As String

Dim XParam As String

Private Sub Acepta_Click()

    Articulo.Text = UCase(Articulo.Text)
    ArticuloDy.Text = UCase(ArticuloDy.Text)

    If Articulo.Text = "  -   -   " Then
        Exit Sub
    End If
    If ArticuloDy.Text = "  -   -   " Then
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Codigo = " + "'" + Articulo.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Exit Sub
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Articulo"
    ZSql = ZSql + " Where Codigo = " + "'" + ArticuloDy.Text + "'"
    spArticulo = ZSql
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Exit Sub
    End If
    




    XParam = "'" + Articulo.Text + "','" _
                 + ArticuloDy.Text + "'"
    
    spComposicion = "ModificaComposicionDy" + XParam
    Set rstComposicion = db.OpenRecordset(spComposicion, dbOpenSnapshot, dbSQLPassThrough)
    
    spCotiza = "ModificaCotizaDy" + XParam
    Set rstCotiza = db.OpenRecordset(spCotiza, dbOpenSnapshot, dbSQLPassThrough)
    
    spEspecificaciones = "ModificaEspecificacionesDy" + XParam
    Set rstEspecificaciones = db.OpenRecordset(spEspecificaciones, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovguia = "ModificaGuiaDy" + XParam
    Set rstMovguia = db.OpenRecordset(spMovguia, dbOpenSnapshot, dbSQLPassThrough)
    
    spHoja = "ModificaHojaDy" + XParam
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
    spInforme = "ModificaInformeDy" + XParam
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    spLaudo = "ModificaLaudoDy" + XParam
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovlab = "ModificaMovlabDy" + XParam
    Set rstMovlab = db.OpenRecordset(spMovlab, dbOpenSnapshot, dbSQLPassThrough)
    
    spMovvar = "ModificaMovvarDy" + XParam
    Set rstMovvar = db.OpenRecordset(spMovvar, dbOpenSnapshot, dbSQLPassThrough)
    
    spOrden = "ModificaOrdenDy" + XParam
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    
    spPrestamo = "ModificaPrestamoDy" + XParam
    Set rstPrestamo = db.OpenRecordset(spPrestamo, dbOpenSnapshot, dbSQLPassThrough)
    
    spPrueart = "ModificaComposicionDy" + XParam
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    
    spArticulo = "ModificaArticuloDy" + XParam
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
    Articulo.Text = "  -   -   "
    ArticuloDy.Text = "  -   -   "
    
    Articulo.SetFocus
    
End Sub

Private Sub Cancela_click()
    PrgPasaDy.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        ArticuloDy.SetFocus
    End If
End Sub

Private Sub ArticuloDy_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ArticuloDy.Text = UCase(ArticuloDy.Text)
        Articulo.SetFocus
    End If
End Sub


Sub Form_Load()
    Articulo.Text = "  -   -   "
    ArticuloDy.Text = "  -   -   "
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


