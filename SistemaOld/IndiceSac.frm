VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form PrgIndiceSac 
   AutoRedraw      =   -1  'True
   Caption         =   "Listado de SAC por Centro"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   ScaleHeight     =   8205
   ScaleWidth      =   15240
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
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   11655
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7680
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   6000
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox Tipo 
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
         Left            =   5160
         TabIndex        =   29
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox Emisor 
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   26
         Text            =   " "
         Top             =   1680
         Width           =   855
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
         Height          =   735
         Left            =   6000
         TabIndex        =   24
         Top             =   1080
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
         Height          =   735
         Left            =   6000
         TabIndex        =   23
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox Centro 
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   20
         Text            =   " "
         Top             =   2400
         Width           =   855
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   17
         Text            =   " "
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox OrdenIII 
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
         Left            =   8640
         TabIndex        =   15
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox OrdenII 
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
         Left            =   5160
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox OrdenI 
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
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox Origen 
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   3495
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
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   3495
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
         Left            =   1560
         MaxLength       =   6
         TabIndex        =   2
         Text            =   " "
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Tipo "
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
         Left            =   3480
         TabIndex        =   30
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label DesEmisor 
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
         Left            =   2520
         TabIndex        =   27
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label DesCentro 
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
         Left            =   2520
         TabIndex        =   22
         Top             =   2400
         Width           =   2535
      End
      Begin VB.Label Label7 
         Caption         =   "Centro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label DesResponsable 
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
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Label Label9 
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Orden Terciario"
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
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Orden Secundario"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Orden Principal"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Pantalla 
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   3480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   7646
      _Version        =   327680
      BackColor       =   16777152
   End
End
Attribute VB_Name = "PrgIndiceSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String
Dim XParam As String
Dim ZZLugar As Integer

Dim ZZTipo(1000, 2) As String
Dim ZZAyudaI(1000) As String
Dim ZZAyudaII(1000) As String
Dim ZZAyudaIII(1000) As String
Dim ZZAyudaIV(1000) As String
Dim ZZAyudaV(1000) As String
Dim ZZAyudaVI(1000) As String

Dim ZVector(1000, 7) As String
Dim ZLugar As Integer

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Emisor.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Emisor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Emisor.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM ResponsableSac"
            ZSql = ZSql + " Where ResponsableSac.Codigo = " + "'" + Emisor.Text + "'"
            spResponsableSac = ZSql
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesEmisor.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Responsable.SetFocus
            End If
                Else
            Responsable.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Emisor.Text = ""
        DesEmisor.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Responsable_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Centro.SetFocus
            End If
                Else
            Centro.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable.Text = ""
        DesResponsable.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Centro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Centro.Text) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CentroSac"
            ZSql = ZSql + " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
            spCentroSac = ZSql
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                DesCentro.Caption = Trim(rstCentroSac!Descripcion)
                rstCentroSac.Close
                Ano.SetFocus
            End If
                Else
            Ano.SetFocus
        End If
        
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
        DesCentro.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Acepta_Click()
    ZZLugar = 3
    Call Opcion
End Sub

Private Sub Cancela_click()
    PrgIndiceSac.Hide
    Unload Me
    Menu.Show
End Sub

Sub Form_Load()

    Erase ZZAyudaI
    Erase ZZAyudaII
    Erase ZZAyudaIII
    Erase ZZAyudaIV
    Erase ZZAyudaV
    Erase ZZAyudaVI
    
    Estado.Clear

    Estado.AddItem "Total"
    Estado.AddItem "INICIADA"
    Estado.AddItem "INVESTIGACION"
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "IMPLEMENTACION A VERIFICAR"
    Estado.AddItem "IMPLEMENTACION VERIFICADA"
    Estado.AddItem "CERRADA"
    
    Estado.ListIndex = 0
    
    ZZAyudaI(1) = "Iniciada"
    ZZAyudaI(2) = "Investig."
    ZZAyudaI(3) = "Implemen."
    ZZAyudaI(4) = "Impl.a Veri"
    ZZAyudaI(5) = "Impl.Verifi"
    ZZAyudaI(6) = "Cerrada"
    
    Origen.Clear
    
    Origen.AddItem "Total"
    Origen.AddItem "Auditoria"
    Origen.AddItem "Reclamo"
    Origen.AddItem "I. No Conformidad"
    Origen.AddItem "Proceso/Sist"
    Origen.AddItem "Otro"
    
    Origen.ListIndex = 0
    
    ZZAyudaII(1) = "Auditoria"
    ZZAyudaII(2) = "Reclamo"
    ZZAyudaII(3) = "I.No Conf"
    ZZAyudaII(4) = "Proc/Sist"
    ZZAyudaII(5) = "Otro"
    
    OrdenI.Clear
    
    OrdenI.AddItem "Tipo"
    OrdenI.AddItem "Numero"
    OrdenI.AddItem "Sector"
    OrdenI.AddItem "Estado"
    OrdenI.AddItem "Responsable"
    
    OrdenI.ListIndex = 0
    
    OrdenII.Clear
    
    OrdenII.AddItem ""
    OrdenII.AddItem "Tipo"
    OrdenII.AddItem "Numero"
    OrdenII.AddItem "Sector"
    OrdenII.AddItem "Estado"
    OrdenII.AddItem "Responsable"
    
    OrdenII.ListIndex = 2
    
    OrdenIII.Clear
    
    OrdenIII.AddItem ""
    OrdenIII.AddItem "Tipo"
    OrdenIII.AddItem "Numero"
    OrdenIII.AddItem "Sector"
    OrdenIII.AddItem "Estado"
    OrdenIII.AddItem "Responsable"
    
    OrdenIII.ListIndex = 0
    
    Erase ZZTipo
    Lugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoSac"
    ZSql = ZSql + " Order by TipoSac.Codigo"
    spTipoSac = ZSql
    Set rstTipoSac = db.OpenRecordset(spTipoSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoSac.RecordCount > 0 Then
        With rstTipoSac
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Lugar = Lugar + 1
                    
                    ZZTipo(Lugar, 1) = rstTipoSac!Codigo
                    ZZTipo(Lugar, 2) = rstTipoSac!Descripcion
                    
                    ZZAyudaIII(rstTipoSac!Codigo) = Trim(rstTipoSac!Descripcion)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstTipoSac.Close
    End If
    
    Tipo.Clear
    
    Tipo.AddItem "Total"
    
    For Ciclo = 1 To Lugar
        Tipo.AddItem Trim(ZZTipo(Ciclo, 2))
    Next Ciclo
    
    Tipo.ListIndex = 0
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CentroSac"
    ZSql = ZSql + " Order by CentroSac.Codigo"
    spCentroSac = ZSql
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        With rstCentroSac
            .MoveFirst
            Do
                If .EOF = False Then
                    ZZAyudaIV(rstCentroSac!Codigo) = Trim(rstCentroSac!Descripcion)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstCentroSac.Close
    End If
    
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
                    ZZAyudaV(rstResponsableSac!Codigo) = Trim(rstResponsableSac!Descripcion)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstResponsableSac.Close
    End If
    
    
    
    
    Ano.Text = Right$(Date$, 4)
    
    Emisor.Text = ""
    DesEmisor.Caption = ""
    Responsable.Text = ""
    DesResponsable.Caption = ""
    Centro.Text = ""
    DesCentro.Caption = ""
    
    ZZLugar = 3
    Call Opcion
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

        Call Limpia_Ayuda
        LugarAyuda = 0
        WIndice.Clear
    
        WEspacios = Len(Ayuda.Text)
    
        Select Case ZZLugar
            Case 1
                Sql1 = "Select *"
                Sql2 = " FROM CentroSac"
                Sql3 = " Order by CentroSac.Codigo"
                spCentroSac = Sql1 + Sql2 + Sql3
                Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCentroSac.RecordCount > 0 Then
                    With rstCentroSac
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                da = Len(rstCentroSac!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstCentroSac!Descripcion, aa, WEspacios) Then
                                        LugarAyuda = LugarAyuda + 1
                                        Pantalla.Row = LugarAyuda
                                        Pantalla.Col = 1
                                        Pantalla.Text = rstCentroSac!Codigo
                                        Pantalla.Col = 2
                                        Pantalla.Text = rstCentroSac!Descripcion
                                        IngresaItem = rstCentroSac!Codigo
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
                    rstCentroSac.Close
                End If
    
            Case 2, 4
                Sql1 = "Select *"
                Sql2 = " FROM ResponsableSac"
                Sql3 = " Order by ResponsableSac.Codigo"
                spResponsableSac = Sql1 + Sql2 + Sql3
                Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstResponsableSac.RecordCount > 0 Then
                    With rstResponsableSac
                        .MoveFirst
                        Do
                            If .EOF = False Then
                                da = Len(rstResponsableSac!Descripcion) - WEspacios
                                For aa = 1 To da + 1
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                        LugarAyuda = LugarAyuda + 1
                                        Pantalla.Row = LugarAyuda
                                        Pantalla.Col = 1
                                        Pantalla.Text = rstResponsableSac!Codigo
                                        Pantalla.Col = 2
                                        Pantalla.Text = rstResponsableSac!Descripcion
                                        IngresaItem = rstResponsableSac!Codigo
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
                    rstResponsableSac.Close
                End If
        End Select
    End If
    
    
    Exit Sub
    
WError:
    Resume Next

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
            Pantalla.Cols = 13
            Pantalla.FixedRows = 1
            Pantalla.Rows = 10001
            
            Pantalla.ColWidth(0) = 200
            Pantalla.Row = 0
            
            For Ciclo = 1 To Pantalla.Cols - 1
                Pantalla.Col = Ciclo
                Select Case Ciclo
                    Case 1
                        Pantalla.Text = "Tipo"
                        Pantalla.ColWidth(Ciclo) = 800
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 2
                        Pantalla.Text = "Año"
                        Pantalla.ColWidth(Ciclo) = 600
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 3
                        Pantalla.Text = "Nro"
                        Pantalla.ColWidth(Ciclo) = 500
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 4
                        Pantalla.Text = "Fecha"
                        Pantalla.ColWidth(Ciclo) = 1200
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 5
                        Pantalla.Text = "Estado"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 6
                        Pantalla.Text = "Titulo"
                        Pantalla.ColWidth(Ciclo) = 4500
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 7
                        Pantalla.Text = "Referencia"
                        Pantalla.ColWidth(Ciclo) = 4500
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 8
                        Pantalla.Text = "Centro"
                        Pantalla.ColWidth(Ciclo) = 1400
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 9
                        Pantalla.Text = "Origen"
                        Pantalla.ColWidth(Ciclo) = 1000
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 10
                        Pantalla.Text = "Emisor"
                        Pantalla.ColWidth(Ciclo) = 800
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 11
                        Pantalla.Text = "Respon."
                        Pantalla.ColWidth(Ciclo) = 800
                        Pantalla.ColAlignment(Ciclo) = flexAlignLeftCenter
                    Case 12
                        Pantalla.Text = ""
                        Pantalla.ColWidth(Ciclo) = 10
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

Private Sub pantalla_Click()
    Indice = Pantalla.Row - 1
    Select Case ZZLugar
        Case 1
            Centro.Text = WIndice.List(Indice)
            Call Centro_Keypress(13)
            Pantalla.Visible = False
        Case 2
            Responsable.Text = WIndice.List(Indice)
            Call Responsable_Keypress(13)
            Pantalla.Visible = False
        Case 3
            WPasaNumero = Pantalla.TextMatrix(Pantalla.Row, 12)
            PrgConsultaSacauto.Show
        
        Case 4
            Emisor.Text = WIndice.List(Indice)
            Call Emisor_Keypress(13)
            Pantalla.Visible = False
        Case Else
    End Select
    Ayuda.Visible = False
End Sub

Private Sub Opcion()

    Rem On Error GoTo WError
    
    Dim IngresaItem As String

    Call Limpia_Ayuda
    LugarAyuda = 0
    WIndice.Clear
    Select Case ZZLugar
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstCentroSac!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstCentroSac!Descripcion
                            IngresaItem = rstCentroSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
            
        Case 2, 4
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            LugarAyuda = LugarAyuda + 1
                            Pantalla.Row = LugarAyuda
                            Pantalla.Col = 1
                            Pantalla.Text = rstResponsableSac!Codigo
                            Pantalla.Col = 2
                            Pantalla.Text = rstResponsableSac!Descripcion
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
            Ayuda.Visible = True
            Ayuda.Text = ""
            Ayuda.SetFocus
            
        Case Else
            If Val(Ano.Text) <> 0 Then
                WDesdeAno = Ano.Text
                WHastaAno = Ano.Text
                    Else
                WDesdeAno = "0"
                WHastaAno = "9999"
            End If
            
            If Val(Responsable.Text) <> 0 Then
                WDesdeRespo = Responsable.Text
                WHastaRespo = Responsable.Text
                    Else
                WDesdeRespo = "0"
                WHastaRespo = "9999"
            End If
            
            If Val(Emisor.Text) <> 0 Then
                WDesdeEemisor = Emisor.Text
                WHastaEmisor = Emisor.Text
                    Else
                WDesdeEmisor = "0"
                WHastaEmisor = "9999"
            End If
            
            If Val(Centro.Text) <> 0 Then
                WDesdeCentro = Centro.Text
                WHastaCentro = Centro.Text
                    Else
                WDesdeCentro = "0"
                WHastaCentro = "9999"
            End If
                    
            Select Case Estado.ListIndex
                Case 0
                    WDesdeEstado = "0"
                    WHastaEstado = "6"
                Case 1
                    WDesdeEstado = "1"
                    WHastaEstado = "1"
                Case 2
                    WDesdeEstado = "2"
                    WHastaEstado = "2"
                Case 3
                    WDesdeEstado = "3"
                    WHastaEstado = "3"
                Case 4
                    WDesdeEstado = "4"
                    WHastaEstado = "4"
                Case 5
                    WDesdeEstado = "5"
                    WHastaEstado = "5"
                Case 6
                    WDesdeEstado = "6"
                    WHastaEstado = "6"
                Case Else
            End Select
            
            Select Case Origen.ListIndex
                Case 0
                    WDesdeOrigen = "0"
                    WHastaOrigen = "5"
                Case 1
                    WDesdeOrigen = "1"
                    WHastaOrigen = "1"
                Case 2
                    WDesdeOrigen = "2"
                    WHastaOrigen = "2"
                Case 3
                    WDesdeOrigen = "3"
                    WHastaOrigen = "3"
                Case 4
                    WDesdeOrigen = "4"
                    WHastaOrigen = "4"
                Case 5
                    WDesdeOrigen = "5"
                    WHastaOrigen = "5"
                Case Else
            End Select
            
            Select Case Tipo.ListIndex
                Case 0
                    WDesdeTipo = "0"
                    WHastaTipo = "9999"
                Case Else
                    WDesdeTipo = ZZTipo(Tipo.ListIndex, 1)
                    WHastaTipo = ZZTipo(Tipo.ListIndex, 1)
            End Select
                    
                        
                        
                        
                        
                    
            ZLugar = 0
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Centro >= " + "'" + WDesdeCentro + "'"
            ZSql = ZSql + " and CargaSac.Centro <= " + "'" + WHastaCentro + "'"
            ZSql = ZSql + " and CargaSac.Estado >= " + "'" + WDesdeEstado + "'"
            ZSql = ZSql + " and CargaSac.Estado <= " + "'" + WHastaEstado + "'"
            ZSql = ZSql + " and CargaSac.Origen >= " + "'" + WDesdeOrigen + "'"
            ZSql = ZSql + " and CargaSac.Origen <= " + "'" + WHastaOrigen + "'"
            ZSql = ZSql + " and CargaSac.Tipo >= " + "'" + WDesdeTipo + "'"
            ZSql = ZSql + " and CargaSac.Tipo <= " + "'" + WHastaTipo + "'"
            ZSql = ZSql + " and CargaSac.ResponsableDestino >= " + "'" + WDesdeRespo + "'"
            ZSql = ZSql + " and CargaSac.ResponsableDestino <= " + "'" + WHastaRespo + "'"
            ZSql = ZSql + " and CargaSac.ResponsableEmisor >= " + "'" + WDesdeEmisor + "'"
            ZSql = ZSql + " and CargaSac.ResponsableEmisor <= " + "'" + WHastaEmisor + "'"
            ZSql = ZSql + " and CargaSac.Ano >= " + "'" + WDesdeAno + "'"
            ZSql = ZSql + " and CargaSac.Ano <= " + "'" + WHastaAno + "'"
            
            Select Case OrdenI.ListIndex
                Case 0
                    ZSql = ZSql + " Order by CargaSac.Tipo"
                Case 1
                    ZSql = ZSql + " Order by CargaSac.Numero"
                Case 2
                    ZSql = ZSql + " Order by CargaSac.Centro"
                Case 3
                    ZSql = ZSql + " Order by CargaSac.Estado"
                Case Else
                    ZSql = ZSql + " Order by CargaSac.ResponsableDestino"
            End Select
            
            Select Case OrdenII.ListIndex
                Case 1
                    ZSql = ZSql + ", CargaSac.Tipo"
                Case 2
                    ZSql = ZSql + ", CargaSac.Numero"
                Case 3
                    ZSql = ZSql + ", CargaSac.Centro"
                Case 4
                    ZSql = ZSql + ", CargaSac.Estado"
                Case 5
                    ZSql = ZSql + ", CargaSac.ResponsableDestino"
                Case Else
            End Select
            
            Select Case OrdenIII.ListIndex
                Case 1
                    ZSql = ZSql + ", CargaSac.Tipo"
                Case 2
                    ZSql = ZSql + ", CargaSac.Numero"
                Case 3
                    ZSql = ZSql + ", CargaSac.Centro"
                Case 4
                    ZSql = ZSql + ", CargaSac.Estado"
                Case 5
                    ZSql = ZSql + ", CargaSac.ResponsableDestino"
                Case Else
            End Select
            
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                With rstCargaSac
                .MoveFirst
                Do
                    If .EOF = False Then
                    
                        ZLugar = ZLugar + 1
                        
                        Pantalla.TextMatrix(ZLugar, 1) = ZZAyudaIII(rstCargaSac!Tipo)
                        Pantalla.TextMatrix(ZLugar, 2) = rstCargaSac!Ano
                        Pantalla.TextMatrix(ZLugar, 3) = rstCargaSac!Numero
                        Pantalla.TextMatrix(ZLugar, 4) = rstCargaSac!Fecha
                        Pantalla.TextMatrix(ZLugar, 5) = ZZAyudaI(rstCargaSac!Estado)
                        Pantalla.TextMatrix(ZLugar, 6) = rstCargaSac!Titulo
                        Pantalla.TextMatrix(ZLugar, 7) = rstCargaSac!Referencia
                        Pantalla.TextMatrix(ZLugar, 8) = ZZAyudaIV(rstCargaSac!Centro)
                        Pantalla.TextMatrix(ZLugar, 9) = ZZAyudaII(rstCargaSac!Origen)
                        Pantalla.TextMatrix(ZLugar, 10) = ZZAyudaV(rstCargaSac!ResponsableEmisor)
                        Pantalla.TextMatrix(ZLugar, 11) = ZZAyudaV(rstCargaSac!ResponsableDestino)
                        Pantalla.TextMatrix(ZLugar, 12) = rstCargaSac!Clave
                        
                        .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCargaSac.Close
            End If
            
    End Select
            
    Pantalla.Visible = True
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Responsable_DblClick()
    ZZLugar = 2
    Call Opcion
End Sub

Private Sub Emisor_DblClick()
    ZZLugar = 4
    Call Opcion
End Sub

Private Sub Centro_DblClick()
    ZZLugar = 1
    Call Opcion
End Sub

