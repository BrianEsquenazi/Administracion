VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgComposicion 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreso de Composicion"
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
   Begin Crystal.CrystalReport Listado 
      Left            =   5400
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wlineas.rpt"
      Destination     =   1
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
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
      ItemData        =   "composicion.frx":0000
      Left            =   360
      List            =   "composicion.frx":0007
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
      Left            =   3000
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
Attribute VB_Name = "PrgComposicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCompo As Recordset
Dim spCompo As String
Dim XParam As String

Private Sub cmdAdd_Click()
    If Codigo.Text <> "" Then
    
        spCompo = "ConsultaCompo " + "'" + Codigo.Text + "'"
        Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
        If rstCompo.RecordCount > 0 Then
            XParam = "'" + Codigo.Text + "','" + Descripcion.Text + "'"
            Set rstCompo = db.OpenRecordset("ModificaCompo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                Else
            XParam = "'" + Codigo.Text + "','" + Descripcion.Text + "'"
            Set rstCompo = db.OpenRecordset("AltaCompo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        End If
        
        Call CmdLimpiar_Click
        Codigo.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
        spCompo = "ConsultaCompo " + "'" + Codigo.Text + "'"
        Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
        If rstCompo.RecordCount > 0 Then
            T$ = "Composicion"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spCompo = "BorrarCompo " + "'" + Codigo.Text + "'"
                Set rstCompo = db.OpenRecordset(spCompo, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = ""
    Descripcion.Text = ""
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Codigo.SetFocus
    PrgComposicion.Hide
    Unload Me
    PrgOt.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
    Codigo.SetFocus
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Codigo.SetFocus
    End If
End Sub

Sub Codigo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Codigo.Text) <> 0 Then
            WCompo = Codigo.Text
            spCompo = "ConsultaCompo " + "'" + Codigo.Text + "'"
            Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
            If rstCompo.RecordCount > 0 Then
                Codigo.Text = rstCompo!Codigo
                Descripcion.Text = rstCompo!Descripcion
                    Else
                WCompo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCompo
            End If
        End If
        Descripcion.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = 0
    
    Select Case XIndice
        Case 0
            spCompo = "ListaCompo"
            Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
            
            With rstCompo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstCompo!Codigo) + " " + rstCompo!Descripcion
                        Pantalla.AddItem IngresaItem
                        IngresaItem = rstCompo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstCompo.Close
        
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
            WCompo = WIndice.List(Indice)
            spCompo = "ConsultaCompo " + "'" + Str$(WCompo) + "'"
            Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
            If rstCompo.RecordCount > 0 Then
                Codigo.Text = rstCompo!Codigo
                Descripcion.Text = rstCompo!Descripcion
                        Else
                CmdLimpiar_Click
                Codigo.Text = WCompo
            End If
            
            Codigo.SetFocus
        
        Case Else
    End Select
    
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spCompo = "ListaCompo"
    Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCompo.RecordCount > 0 Then
        With rstCompo
            .MoveFirst
            Codigo.Text = rstCompo!Codigo
            Descripcion.Text = rstCompo!Descripcion
        End With
        rstCompo.Close
    End If
    Codigo.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Composicion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spCompo = "ListaCompo"
    Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCompo.RecordCount > 0 Then
        With rstCompo
            .MoveLast
            Codigo.Text = rstCompo!Codigo
            Descripcion.Text = rstCompo!Descripcion
            Codigo.SetFocus
        End With
        rstCompo.Close
    End If
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Composicion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    spCompo = "AnteriorCompo " + "'" + Codigo.Text + "'"
    Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCompo.RecordCount > 0 Then
        With rstCompo
            .MoveLast
            Codigo.Text = rstCompo!Codigo
            Descripcion.Text = rstCompo!Descripcion
        End With
        rstCompo.Close
    End If
    
    Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Composicion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
    
End Sub


Private Sub Siguiente_Click()

    On Error GoTo WError
    
    spCompo = "PosteriorCompo " + "'" + Codigo.Text + "'"
    Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCompo.RecordCount > 0 Then
        With rstCompo
            .MoveFirst
            Codigo.Text = rstCompo!Codigo
            Descripcion.Text = rstCompo!Descripcion
        End With
        rstCompo.Close
    End If
    
    Codigo.SetFocus
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Composicion", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Codigo.SetFocus
    
End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    spCompo = "ListaCompo"
    Set rstCompo = db.OpenRecordset(spCompo, dbOpenSnapshot, dbSQLPassThrough)
    If rstCompo.RecordCount > 0 Then
        With rstCompo
            .MoveFirst
            Do
                If .EOF = False Then
            
                    da = Len(rstCompo!Descripcion) - WEspacios
                
                    For aa = 1 To da
                        If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                            IngresaItem = Str$(rstCompo!Codigo) + " " + rstCompo!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCompo!Codigo
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
        rstCompo.Close
    End If
    End If

End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgComposicion.Caption = "Ingreso de Composicion :  " + !Nombre
        End If
    End With
End Sub


