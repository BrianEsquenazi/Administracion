VERSION 5.00
Begin VB.Form PrgOrdenComplementoImpo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de datos Complementarios de la Orden de Compra"
   ClientHeight    =   8115
   ClientLeft      =   180
   ClientTop       =   375
   ClientWidth     =   11685
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8115
   ScaleWidth      =   11685
   Visible         =   0   'False
   Begin VB.TextBox Descri33 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   9
      Top             =   2520
      Width           =   8415
   End
   Begin VB.TextBox Descri14 
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
      Left            =   5520
      MaxLength       =   70
      TabIndex        =   3
      Top             =   720
      Width           =   6015
   End
   Begin VB.CommandButton FinIngresaObserva 
      Caption         =   "Fin de Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   30
      Top             =   7440
      Width           =   2175
   End
   Begin VB.TextBox Descri56 
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
      TabIndex        =   25
      Top             =   6720
      Width           =   8415
   End
   Begin VB.TextBox Descri57 
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
      TabIndex        =   26
      Top             =   7080
      Width           =   8415
   End
   Begin VB.TextBox Descri54 
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
      TabIndex        =   23
      Top             =   6240
      Width           =   8415
   End
   Begin VB.TextBox Descri55 
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
      TabIndex        =   24
      Top             =   6480
      Width           =   8415
   End
   Begin VB.TextBox Descri51 
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
      TabIndex        =   20
      Top             =   5400
      Width           =   8415
   End
   Begin VB.TextBox Descri52 
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
      TabIndex        =   21
      Top             =   5640
      Width           =   8415
   End
   Begin VB.TextBox Descri53 
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
      TabIndex        =   22
      Top             =   5880
      Width           =   8415
   End
   Begin VB.TextBox Descri42 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   12
      Top             =   3360
      Width           =   8415
   End
   Begin VB.TextBox Descri41 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   11
      Top             =   3120
      Width           =   8415
   End
   Begin VB.TextBox Descri40 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   10
      Top             =   2880
      Width           =   8415
   End
   Begin VB.TextBox Descri44 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   14
      Top             =   3840
      Width           =   8415
   End
   Begin VB.TextBox Descri43 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   13
      Top             =   3600
      Width           =   8415
   End
   Begin VB.TextBox Descri47 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   17
      Top             =   4560
      Width           =   8415
   End
   Begin VB.TextBox Descri46 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   16
      Top             =   4320
      Width           =   8415
   End
   Begin VB.TextBox Descri45 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   15
      Top             =   4080
      Width           =   8415
   End
   Begin VB.TextBox Descri48 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   18
      Top             =   4800
      Width           =   8415
   End
   Begin VB.TextBox Descri49 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   19
      Top             =   5040
      Width           =   8415
   End
   Begin VB.TextBox Descri31 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   7
      Top             =   2040
      Width           =   8415
   End
   Begin VB.TextBox Descri32 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   8
      Top             =   2280
      Width           =   8415
   End
   Begin VB.TextBox Descri21 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   4
      Top             =   1200
      Width           =   8415
   End
   Begin VB.TextBox Descri22 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   5
      Top             =   1440
      Width           =   8415
   End
   Begin VB.TextBox Descri23 
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
      Left            =   120
      MaxLength       =   70
      TabIndex        =   6
      Top             =   1680
      Width           =   8415
   End
   Begin VB.TextBox Descri11 
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
      Left            =   5520
      MaxLength       =   70
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.TextBox Descri12 
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
      Left            =   5520
      MaxLength       =   70
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
   Begin VB.TextBox Descri13 
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
      Left            =   5520
      MaxLength       =   70
      TabIndex        =   2
      Top             =   480
      Width           =   6015
   End
   Begin VB.Label Label6 
      Caption         =   "Nro de Cuenta"
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
      TabIndex        =   33
      Top             =   7080
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Copias"
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
      TabIndex        =   32
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Originales"
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
      TabIndex        =   31
      Top             =   5400
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Incoterms"
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
      Left            =   3120
      TabIndex        =   29
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Condiciones de Pago"
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
      Left            =   3120
      TabIndex        =   28
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Instrucciones de Envio"
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
      Left            =   3120
      TabIndex        =   27
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "PrgOrdenComplementoImpo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstObservaOrdenImpo As Recordset
Dim spObservaOrdenImpo As String
Dim rstProveedorAdicional As Recordset
Dim spProveedorAdicional As String

Dim EmpresaAnterior As String
Dim EmpresaOrden As Integer

Dim WControl As String

Private Sub cmdClose_Click()
    PrgOrdenComplementoImpo.Hide
    Unload Me
    PrgOrden.Show
End Sub

Private Sub Descri11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri12.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri11.Text = ""
    End If
End Sub

Private Sub Descri12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri13.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri12.Text = ""
    End If
End Sub

Private Sub Descri13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri14.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri13.Text = ""
    End If
End Sub

Private Sub Descri14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri21.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri14.Text = ""
    End If
End Sub

Private Sub Descri21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri22.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri21.Text = ""
    End If
End Sub

Private Sub Descri22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri23.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri22.Text = ""
    End If
End Sub

Private Sub Descri23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri31.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri23.Text = ""
    End If
End Sub

Private Sub Descri31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri32.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri31.Text = ""
    End If
End Sub

Private Sub Descri32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri33.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri32.Text = ""
    End If
End Sub

Private Sub Descri33_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri40.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri33.Text = ""
    End If
End Sub

Private Sub Descri40_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri41.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri40.Text = ""
    End If
End Sub

Private Sub Descri41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri42.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri41.Text = ""
    End If
End Sub

Private Sub Descri42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri43.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri42.Text = ""
    End If
End Sub

Private Sub Descri43_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri44.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri43.Text = ""
    End If
End Sub

Private Sub Descri44_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri45.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri44.Text = ""
    End If
End Sub

Private Sub Descri45_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri46.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri45.Text = ""
    End If
End Sub

Private Sub Descri46_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri47.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri46.Text = ""
    End If
End Sub

Private Sub Descri47_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri48.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri47.Text = ""
    End If
End Sub

Private Sub Descri48_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri49.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri48.Text = ""
    End If
End Sub

Private Sub Descri49_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri51.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri49.Text = ""
    End If
End Sub

Private Sub Descri51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri52.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri51.Text = ""
    End If
End Sub

Private Sub Descri52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri53.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri52.Text = ""
    End If
End Sub

Private Sub Descri53_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri54.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri53.Text = ""
    End If
End Sub

Private Sub Descri54_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri55.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri54.Text = ""
    End If
End Sub

Private Sub Descri55_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri56.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri55.Text = ""
    End If
End Sub

Private Sub Descri56_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri57.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri56.Text = ""
    End If
End Sub

Private Sub Descri57_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descri11.SetFocus
    End If
    If KeyAscii = 27 Then
        Descri57.Text = ""
    End If
End Sub

Private Sub Form_Load()

    ZOrden = WPasaOrden
    ZCarpeta = WPasaCarpeta
    ZProveedor = WPasaProveedor
    
    Descri11.Text = ""
    Descri12.Text = ""
    Descri13.Text = ""
    Descri14.Text = ""
    
    Descri21.Text = ""
    Descri22.Text = ""
    Descri23.Text = ""
    
    Descri31.Text = ""
    Descri32.Text = ""
    Descri33.Text = ""
    
    Descri40.Text = ""
    Descri41.Text = ""
    Descri42.Text = ""
    Descri43.Text = ""
    Descri44.Text = ""
    Descri45.Text = ""
    Descri46.Text = ""
    Descri47.Text = ""
    Descri48.Text = ""
    Descri49.Text = ""
    
    Descri51.Text = ""
    Descri52.Text = ""
    Descri53.Text = ""
    Descri54.Text = ""
    Descri55.Text = ""
    Descri56.Text = ""
    Descri57.Text = ""
    
    Lee = "S"
    
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrdenImpo"
    Sql3 = " Where ObservaOrdenImpo.Orden = " + "'" + ZOrden + "'"
    spObservaOrdenImpo = Sql1 + Sql2 + Sql3
    Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrdenImpo.RecordCount > 0 Then
    
        Lee = "N"
        
        Descri11.Text = rstObservaOrdenImpo!Descri11
        Descri12.Text = rstObservaOrdenImpo!Descri12
        Descri13.Text = rstObservaOrdenImpo!Descri13
        Descri14.Text = rstObservaOrdenImpo!Descri14
        
        Descri21.Text = rstObservaOrdenImpo!Descri21
        Descri22.Text = rstObservaOrdenImpo!Descri22
        Descri23.Text = rstObservaOrdenImpo!Descri23
        
        Descri31.Text = rstObservaOrdenImpo!Descri31
        Descri32.Text = rstObservaOrdenImpo!Descri32
        Descri33.Text = rstObservaOrdenImpo!Descri33
        
        Descri40.Text = rstObservaOrdenImpo!Descri40
        Descri41.Text = rstObservaOrdenImpo!Descri41
        Descri42.Text = rstObservaOrdenImpo!Descri42
        Descri43.Text = rstObservaOrdenImpo!Descri43
        Descri44.Text = rstObservaOrdenImpo!Descri44
        Descri45.Text = rstObservaOrdenImpo!Descri45
        Descri46.Text = rstObservaOrdenImpo!Descri46
        Descri47.Text = rstObservaOrdenImpo!Descri47
        Descri48.Text = rstObservaOrdenImpo!Descri48
        Descri49.Text = rstObservaOrdenImpo!Descri49
        
        Descri51.Text = rstObservaOrdenImpo!Descri51
        Descri52.Text = rstObservaOrdenImpo!Descri52
        Descri53.Text = rstObservaOrdenImpo!Descri53
        Descri54.Text = rstObservaOrdenImpo!Descri54
        Descri55.Text = rstObservaOrdenImpo!Descri55
        Descri56.Text = rstObservaOrdenImpo!Descri56
        Descri57.Text = rstObservaOrdenImpo!Descri57
    
        rstObservaOrdenImpo.Close
    End If
    Rem BY NAN
    
    If Lee = "S" Then
        
        XEmpresa = WEmpresa
        
        WEmpresa = "0001"
        txtOdbc = "Empresa01"
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        Sql1 = "Select *"
        Sql2 = " FROM ProveedorAdicional"
        Sql3 = " Where ProveedorAdicional.Proveedor = " + "'" + ZProveedor + "'"
        spProveedorAdicional = Sql1 + Sql2 + Sql3 + Sql4
        Set rstProveedorAdicional = db.OpenRecordset(spProveedorAdicional, dbOpenSnapshot, dbSQLPassThrough)
        If rstProveedorAdicional.RecordCount > 0 Then
            
            Descri11.Text = rstProveedorAdicional!Descri11
            Descri12.Text = rstProveedorAdicional!Descri12
            Descri13.Text = rstProveedorAdicional!Descri13
            Descri14.Text = rstProveedorAdicional!Descri14
            
            Descri21.Text = rstProveedorAdicional!Descri21
            Descri22.Text = rstProveedorAdicional!Descri22
            Descri23.Text = rstProveedorAdicional!Descri23
            
            Descri31.Text = rstProveedorAdicional!Descri31
            Descri32.Text = rstProveedorAdicional!Descri32
            Descri33.Text = rstProveedorAdicional!Descri33
            
            Descri40.Text = rstProveedorAdicional!Descri40
            Descri41.Text = rstProveedorAdicional!Descri41
            Descri42.Text = rstProveedorAdicional!Descri42
            Descri43.Text = rstProveedorAdicional!Descri43
            Descri44.Text = rstProveedorAdicional!Descri44
            Descri45.Text = rstProveedorAdicional!Descri45
            Descri46.Text = rstProveedorAdicional!Descri46
            Descri47.Text = rstProveedorAdicional!Descri47
            Descri48.Text = rstProveedorAdicional!Descri48
            Descri49.Text = rstProveedorAdicional!Descri49
            
            Descri51.Text = rstProveedorAdicional!Descri51
            Descri52.Text = rstProveedorAdicional!Descri52
            Descri53.Text = rstProveedorAdicional!Descri53
            Descri54.Text = rstProveedorAdicional!Descri54
            Descri55.Text = rstProveedorAdicional!Descri55
            Descri56.Text = rstProveedorAdicional!Descri56
            Descri57.Text = rstProveedorAdicional!Descri57
        
            rstProveedorAdicional.Close
        End If
    
        Call Conecta_Empresa
        
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub




Private Sub FinIngresaObserva_Click()

    ZDescri11 = Descri11.Text
    ZDescri12 = Descri12.Text
    ZDescri13 = Descri13.Text
    ZDescri14 = Descri14.Text
    ZDescri21 = Descri21.Text
    ZDescri22 = Descri22.Text
    ZDescri23 = Descri23.Text
    ZDescri31 = Descri31.Text
    ZDescri32 = Descri32.Text
    ZDescri33 = Descri33.Text
    ZDescri40 = Descri40.Text
    ZDescri41 = Descri41.Text
    ZDescri42 = Descri42.Text
    ZDescri43 = Descri43.Text
    ZDescri44 = Descri44.Text
    ZDescri45 = Descri45.Text
    ZDescri46 = Descri46.Text
    ZDescri47 = Descri47.Text
    ZDescri48 = Descri48.Text
    ZDescri49 = Descri49.Text
    ZDescri51 = Descri51.Text
    ZDescri52 = Descri52.Text
    ZDescri53 = Descri53.Text
    ZDescri54 = Descri54.Text
    ZDescri55 = Descri55.Text
    ZDescri56 = Descri56.Text
    ZDescri57 = Descri57.Text
    
    Sql1 = "Select *"
    Sql2 = " FROM ObservaOrdenImpo"
    Sql3 = " Where ObservaOrdenImpo.Orden = " + "'" + WPasaOrden + "'"
    spObservaOrdenImpo = Sql1 + Sql2 + Sql3
    Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
    If rstObservaOrdenImpo.RecordCount > 0 Then
        rstObservaOrdenImpo.Close
        
        ZSql = ""
        ZSql = ZSql + "UPDATE ObservaOrdenImpo SET "
        ZSql = ZSql + " Descri11 = " + "'" + Descri11.Text + "',"
        ZSql = ZSql + " Descri12 = " + "'" + Descri12.Text + "',"
        ZSql = ZSql + " Descri13 = " + "'" + Descri13.Text + "',"
        ZSql = ZSql + " Descri14 = " + "'" + Descri14.Text + "',"
        ZSql = ZSql + " Descri21 = " + "'" + Descri21.Text + "',"
        ZSql = ZSql + " Descri22 = " + "'" + Descri22.Text + "',"
        ZSql = ZSql + " Descri23 = " + "'" + Descri23.Text + "',"
        ZSql = ZSql + " Descri31 = " + "'" + Descri31.Text + "',"
        ZSql = ZSql + " Descri32 = " + "'" + Descri32.Text + "',"
        ZSql = ZSql + " Descri33 = " + "'" + Descri33.Text + "',"
        ZSql = ZSql + " Descri40 = " + "'" + Descri40.Text + "',"
        ZSql = ZSql + " Descri41 = " + "'" + Descri41.Text + "',"
        ZSql = ZSql + " Descri42 = " + "'" + Descri42.Text + "',"
        ZSql = ZSql + " Descri43 = " + "'" + Descri43.Text + "',"
        ZSql = ZSql + " Descri44 = " + "'" + Descri44.Text + "',"
        ZSql = ZSql + " Descri45 = " + "'" + Descri45.Text + "',"
        ZSql = ZSql + " Descri46 = " + "'" + Descri46.Text + "',"
        ZSql = ZSql + " Descri47 = " + "'" + Descri47.Text + "',"
        ZSql = ZSql + " Descri48 = " + "'" + Descri48.Text + "',"
        ZSql = ZSql + " Descri49 = " + "'" + Descri49.Text + "',"
        ZSql = ZSql + " Descri51 = " + "'" + Descri51.Text + "',"
        ZSql = ZSql + " Descri52 = " + "'" + Descri52.Text + "',"
        ZSql = ZSql + " Descri53 = " + "'" + Descri53.Text + "',"
        ZSql = ZSql + " Descri54 = " + "'" + Descri54.Text + "',"
        ZSql = ZSql + " Descri55 = " + "'" + Descri55.Text + "',"
        ZSql = ZSql + " Descri56 = " + "'" + Descri56.Text + "',"
        ZSql = ZSql + " Descri57 = " + "'" + Descri57.Text + "'"
        ZSql = ZSql + " Where Orden = " + "'" + WPasaOrden + "'"
        spObservaOrdenImpo = ZSql
        Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ObservaOrdenImpo ("
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Descri11 ,"
        ZSql = ZSql + "Descri12 ,"
        ZSql = ZSql + "Descri13 ,"
        ZSql = ZSql + "Descri14 ,"
        ZSql = ZSql + "Descri21 ,"
        ZSql = ZSql + "Descri22 ,"
        ZSql = ZSql + "Descri23 ,"
        ZSql = ZSql + "Descri31 ,"
        ZSql = ZSql + "Descri32 ,"
        ZSql = ZSql + "Descri33 ,"
        ZSql = ZSql + "Descri40 ,"
        ZSql = ZSql + "Descri41 ,"
        ZSql = ZSql + "Descri42 ,"
        ZSql = ZSql + "Descri43 ,"
        ZSql = ZSql + "Descri44 ,"
        ZSql = ZSql + "Descri45 ,"
        ZSql = ZSql + "Descri46 ,"
        ZSql = ZSql + "Descri47 ,"
        ZSql = ZSql + "Descri48 ,"
        ZSql = ZSql + "Descri49 ,"
        ZSql = ZSql + "Descri51 ,"
        ZSql = ZSql + "Descri52 ,"
        ZSql = ZSql + "Descri53 ,"
        ZSql = ZSql + "Descri54 ,"
        ZSql = ZSql + "Descri55 ,"
        ZSql = ZSql + "Descri56 ,"
        ZSql = ZSql + "Descri57 )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WPasaOrden + "',"
        ZSql = ZSql + "'" + ZDescri11 + "',"
        ZSql = ZSql + "'" + ZDescri12 + "',"
        ZSql = ZSql + "'" + ZDescri13 + "',"
        ZSql = ZSql + "'" + ZDescri14 + "',"
        ZSql = ZSql + "'" + ZDescri21 + "',"
        ZSql = ZSql + "'" + ZDescri22 + "',"
        ZSql = ZSql + "'" + ZDescri23 + "',"
        ZSql = ZSql + "'" + ZDescri31 + "',"
        ZSql = ZSql + "'" + ZDescri32 + "',"
        ZSql = ZSql + "'" + ZDescri33 + "',"
        ZSql = ZSql + "'" + ZDescri40 + "',"
        ZSql = ZSql + "'" + ZDescri41 + "',"
        ZSql = ZSql + "'" + ZDescri42 + "',"
        ZSql = ZSql + "'" + ZDescri43 + "',"
        ZSql = ZSql + "'" + ZDescri44 + "',"
        ZSql = ZSql + "'" + ZDescri45 + "',"
        ZSql = ZSql + "'" + ZDescri46 + "',"
        ZSql = ZSql + "'" + ZDescri47 + "',"
        ZSql = ZSql + "'" + ZDescri48 + "',"
        ZSql = ZSql + "'" + ZDescri49 + "',"
        ZSql = ZSql + "'" + ZDescri51 + "',"
        ZSql = ZSql + "'" + ZDescri52 + "',"
        ZSql = ZSql + "'" + ZDescri53 + "',"
        ZSql = ZSql + "'" + ZDescri54 + "',"
        ZSql = ZSql + "'" + ZDescri55 + "',"
        ZSql = ZSql + "'" + ZDescri56 + "',"
        ZSql = ZSql + "'" + ZDescri57 + "')"
        spObservaOrdenImpo = ZSql
        Set rstObservaOrdenImpo = db.OpenRecordset(spObservaOrdenImpo, dbOpenSnapshot, dbSQLPassThrough)
        
    End If

    Select Case WPasaOrigen
        Case 1
            PrgOrdenComplementoImpo.Hide
            Unload Me
            PrgOrden.Show
        Case 2
            PrgOrdenComplementoImpo.Hide
            Unload Me
            PrgMovgas.Show
        Case 3
            PrgOrdenComplementoImpo.Hide
            Unload Me
            PrgOrdenImpo.Show
        Case Else
            Close
            End
    End Select
    
End Sub
