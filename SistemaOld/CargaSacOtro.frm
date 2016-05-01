VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgCargaSacOtro 
   Caption         =   "Carga de SAC"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   2265
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   11775
   Begin VB.TextBox Numero 
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
      Left            =   3360
      MaxLength       =   6
      TabIndex        =   2
      Text            =   " "
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Ano 
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
      MaxLength       =   6
      TabIndex        =   1
      Text            =   " "
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Tipo 
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
      Left            =   600
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox IngresoCausa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3240
      Width           =   11415
   End
   Begin VB.TextBox IngresoNoCon 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   480
      Width           =   11415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Descripcion de la No Conformidad"
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
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   11415
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Causas que lo originaron"
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
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   11295
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5640
      MouseIcon       =   "CargaSacOtro.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacOtro.frx":030A
      ToolTipText     =   "Salida"
      Top             =   5520
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   4800
      MouseIcon       =   "CargaSacOtro.frx":0B4C
      MousePointer    =   99  'Custom
      Picture         =   "CargaSacOtro.frx":0E56
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   5520
      Width           =   480
   End
End
Attribute VB_Name = "PrgCargaSacOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoSac As Recordset
Dim spTipoSac As String
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim XParam As String
Dim ZZLugar As Integer

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM TipoSac"
    Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
    spTipoSac = Sql1 + Sql2 + Sql3
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        DesTipo.Caption = Trim(rstTipoSac!Descripcion)
        rstTipoSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableEmisor.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableEmisor.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + ResponsableDestino.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsableDestino.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If

End Sub

Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZTipo = Tipo.Text
    ZAno = Ano.Text
    ZNumero = Numero.Text
    
    Rem Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Tipo.Text = ZTipo
    Ano.Text = ZAno
    Numero.Text = ZNumero
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
    ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
    
        Centro.Text = rstCargaSac!Centro
        Fecha.Text = rstCargaSac!Fecha
        Origen.ListIndex = rstCargaSac!Origen
        Estado.ListIndex = rstCargaSac!Estado
        ResponsableEmisor.Text = rstCargaSac!ResponsableEmisor
        ResponsableDestino.Text = rstCargaSac!ResponsableDestino
        Referencia.Text = Trim(rstCargaSac!Referencia)
        Titulo.Text = Trim(rstCargaSac!Titulo)
        IngresoNoCon.Text = IIf(IsNull(rstCargaSac!IngresoNoCon), "", rstCargaSac!IngresoNoCon)
        IngresoCausa.Text = IIf(IsNull(rstCargaSac!IngresoCausa), "", rstCargaSac!IngresoCausa)
        
        rstCargaSac.Close
    End If
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub cmdAdd_Click()

    If Tipo.Text <> "" And Ano.Text <> "" And Numero.Text <> "" Then
        
        Auxi3 = Tipo.Text
        Auxi1 = Ano.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi3, 4)
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi3 + Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            rstCargaSac.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSac SET "
            ZSql = ZSql + " IngresoCausa = " + "'" + IngresoCausa.Text + "'"
            ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
        ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            ZZEstado = rstCargaSac!Estado
            rstCargaSac.Close
            
            If ZZEstado <= 1 And Trim(IngresoCausa.Text) <> "" Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE CargaSac SET "
                ZSql = ZSql + " Estado = " + "'" + "2" + "'"
                ZSql = ZSql + " Where Tipo = " + "'" + Tipo.Text + "'"
                ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
                ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
        End If
        
        Call cmdClose_Click
        
    End If
End Sub

Private Sub cmdClose_Click()
    PrgCargaSacOtro.Hide
    Unload Me
    PrgCargaSacAccion.Show
End Sub

Sub Form_Load()

    Tipo.Text = WPasaTipo
    Ano.Text = WPasaAno
    Numero.Text = WPasaNumero
    
    Call Numero_KeyPress(13)
    
End Sub


Private Sub Tipo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM TipoSac"
        Sql3 = " Where TipoSac.Codigo = " + "'" + Tipo.Text + "'"
        spTipoSac = Sql1 + Sql2 + Sql3
        Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoSac.RecordCount > 0 Then
            DesTipo.Caption = Trim(rstTipoSac!Descripcion)
            rstTipoSac.Close
            Ano.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Tipo.Text = ""
        DesTipo.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Numero.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Tipo = " + "'" + Tipo.Text + "'"
            ZSql = ZSql + " and CargaSac.Ano = " + "'" + Ano.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                IngresoNoCon.Text = rstCargaSac!IngresoNoCon
                IngresoCausa.Text = rstCargaSac!IngresoCausa
                rstCargaSac.Close
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

