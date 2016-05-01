VERSION 5.00
Begin VB.Form PrgTipoPrv 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Tipos de Proveedores"
   ClientHeight    =   5055
   ClientLeft      =   2730
   ClientTop       =   1425
   ClientWidth     =   6720
   LinkTopic       =   "Form2"
   ScaleHeight     =   5055
   ScaleWidth      =   6720
   Begin VB.TextBox Ayuda 
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
      Left            =   360
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   5775
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
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   15
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5400
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox Pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "TipoPrv.frx":0000
      Left            =   360
      List            =   "TipoPrv.frx":0007
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
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
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
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
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
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
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   4
      Top             =   960
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
      Left            =   2520
      MaxLength       =   50
      TabIndex        =   3
      Top             =   480
      Width           =   3375
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
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
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
      Left            =   360
      TabIndex        =   1
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "PrgTipoPrv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstTipoProv As Recordset
Dim spTipoProv As String
Dim XParam As String

Private Sub Acepta_Click()

    LISTADO.WindowTitle = "Listado de Tipos de Proveedores"
    LISTADO.WindowTop = 0
    LISTADO.WindowLeft = 0
    LISTADO.WindowWidth = Screen.Width
    LISTADO.WindowHeight = Screen.Height

    LISTADO.GroupSelectionFormula = "{TipoProv.Codigo} in " + Desde.Text + " to " + Hasta.Text
    If Impresora.Value = True Then
        LISTADO.Destination = 1
            Else
        LISTADO.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    LISTADO.SQLQuery = "SELECT"
    
    LISTADO.DataFiles(1) = WEmpresa + "auxi.mdb"
    LISTADO.Connect = Connect()
    
    Codigo.SetFocus
    LISTADO.Action = 1
    Frame2.Visible = False
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoProv"
        ZSql = ZSql + " Where TipoProv.Codigo = " + "'" + Codigo.Text + "'"
        spTipoProv = ZSql
        Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoProv.RecordCount > 0 Then
            rstTipoProv.Close
            ZSql = ""
            ZSql = ZSql + "UPDATE TipoProv SET "
            ZSql = ZSql + " Descripcion = " + "'" + Descripcion.Text + "'"
            ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
            spTipoProv = ZSql
            Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
                Else
            ZSql = ""
            ZSql = ZSql + "INSERT INTO TipoProv ("
            ZSql = ZSql + "Codigo ,"
            ZSql = ZSql + "Descripcion )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Descripcion.Text + "')"
            spTipoProv = ZSql
            Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Codigo.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM TipoProv"
        ZSql = ZSql + " Where TipoProv.Codigo = " + "'" + Codigo.Text + "'"
        spTipoProv = ZSql
        Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
        If rstTipoProv.RecordCount > 0 Then
            rstTipoProv.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            TipoProv% = MsgBox(m$, 32 + 4, T$)
            If TipoProv% = 6 Then
                ZSql = ""
                ZSql = ZSql + "DELETE TipoProv"
                ZSql = ZSql + " Where Codigo = " + "'" + Codigo.Text + "'"
                spTipoProv = ZSql
                Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = ""
    Descripcion.Text = ""
    
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM TipoProv"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        rstTipoProv.MoveLast
        ZUltimo = IIf(IsNull(rstTipoProv!CodigoMayor), "0", rstTipoProv!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstTipoProv.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
    End If
End Sub

Private Sub cmdClose_Click()
    PrgTipoPrv.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()
    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM TipoProv"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        rstTipoProv.MoveLast
        ZUltimo = IIf(IsNull(rstTipoProv!CodigoMayor), "0", rstTipoProv!CodigoMayor)
        Codigo.Text = ZUltimo + 1
        rstTipoProv.Close
    End If
    If Val(Codigo.Text) = 0 Then
        Codigo.Text = "1"
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

Sub Desde_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Hasta_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
    Codigo.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Codigo.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion.Text = ""
    End If
End Sub

Sub Codigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            WCodigo = Codigo.Text
            
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoProv"
            ZSql = ZSql + " Where TipoProv.Codigo = " + "'" + Codigo.Text + "'"
            spTipoProv = ZSql
            Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoProv.RecordCount > 0 Then
                Descripcion.Text = rstTipoProv!Descripcion
                rstTipoProv.Close
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

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoProv"
            ZSql = ZSql + " Order by Codigo"
            spTipoProv = ZSql
            Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoProv.RecordCount > 0 Then
                With rstTipoProv
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstTipoProv!Codigo) + " " + rstTipoProv!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstTipoProv!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTipoProv.Close
            End If
            
        Case Else
    End Select
            
    Ayuda.Text = ""
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_Keypress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    ZSql = ""
    ZSql = ZSql + "Select Min(Codigo) as [CodigoMenor]"
    ZSql = ZSql + " FROM TipoProv"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        rstTipoProv.MoveFirst
        ZUltimo = IIf(IsNull(rstTipoProv!CodigoMenor), "0", rstTipoProv!CodigoMenor)
        Codigo.Text = ZUltimo
        rstTipoProv.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoProv"
    ZSql = ZSql + " Where TipoProv.Codigo = " + "'" + Codigo.Text + "'"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        Descripcion.Text = rstTipoProv!Descripcion
        rstTipoProv.Close
    End If
    
    Codigo.SetFocus
    
 End Sub

Private Sub Ultimo_Click()

    ZSql = ""
    ZSql = ZSql + "Select Max(Codigo) as [CodigoMayor]"
    ZSql = ZSql + " FROM TipoProv"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        rstTipoProv.MoveLast
        ZUltimo = IIf(IsNull(rstTipoProv!CodigoMayor), "0", rstTipoProv!CodigoMayor)
        Codigo.Text = ZUltimo
        rstTipoProv.Close
    End If
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoProv"
    ZSql = ZSql + " Where TipoProv.Codigo = " + "'" + Codigo.Text + "'"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        Descripcion.Text = rstTipoProv!Descripcion
        rstTipoProv.Close
    End If
    
    Codigo.SetFocus
    
 End Sub

Private Sub Siguiente_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoProv"
    ZSql = ZSql + " Where TipoProv.Codigo > " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by TipoProv.Codigo"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        With rstTipoProv
            .MoveFirst
            Codigo.Text = rstTipoProv!Codigo
            Descripcion.Text = rstTipoProv!Descripcion
        End With
        rstTipoProv.Close
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Archivo de Tipo de Proveedores")
    End If

End Sub

Private Sub Anterior_Click()

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM TipoProv"
    ZSql = ZSql + " Where TipoProv.Codigo < " + "'" + Codigo.Text + "'"
    ZSql = ZSql + " Order by TipoProv.Codigo"
    spTipoProv = ZSql
    Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
    If rstTipoProv.RecordCount > 0 Then
        With rstTipoProv
            .MoveLast
            Codigo.Text = rstTipoProv!Codigo
            Descripcion.Text = rstTipoProv!Descripcion
        End With
        rstTipoProv.Close
        Codigo.SetFocus
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Archivo de Tipo de Proveedores")
    End If
    
End Sub


Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM TipoProv"
            ZSql = ZSql + " Order by Codigo"
            spTipoProv = ZSql
            Set rstTipoProv = db.OpenRecordset(spTipoProv, dbOpenSnapshot, dbSQLPassThrough)
            If rstTipoProv.RecordCount > 0 Then
                With rstTipoProv
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstTipoProv!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(rstTipoProv!Descripcion), aa, WEspacios) Then
                                    IngresaItem = Str$(rstTipoProv!Codigo) + " " + rstTipoProv!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTipoProv!Codigo
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
                rstTipoProv.Close
            End If
            
        Case Else
    End Select
    
    End If
    
    If KeyAscii = 27 Then
        Ayuda.Text = ""
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgTipoPrv.Caption = "Ingreso de Tipos de Proveedores :  " + !Nombre
        End If
    End With
End Sub


