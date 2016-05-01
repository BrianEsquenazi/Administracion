VERSION 5.00
Begin VB.Form LEYENDAS 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "ABM de Leyendas"
   ClientHeight    =   3990
   ClientLeft      =   1050
   ClientTop       =   1170
   ClientWidth     =   7365
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3990
   ScaleWidth      =   7365
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Panel3D4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5385
      ScaleWidth      =   9225
      TabIndex        =   8
      Top             =   1200
      Width           =   9255
      Begin VB.PictureBox Frame3D1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   2295
         Left            =   600
         ScaleHeight     =   2265
         ScaleWidth      =   7905
         TabIndex        =   18
         Top             =   2880
         Width           =   7935
         Begin VB.PictureBox Command3D7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5160
            ScaleHeight     =   465
            ScaleWidth      =   2385
            TabIndex        =   21
            Top             =   1560
            Width           =   2415
         End
         Begin VB.PictureBox Command3D10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6360
            ScaleHeight     =   465
            ScaleWidth      =   1185
            TabIndex        =   20
            Top             =   960
            Width           =   1215
         End
         Begin VB.PictureBox Command3D8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2520
            ScaleHeight     =   465
            ScaleWidth      =   2145
            TabIndex        =   0
            Top             =   1560
            Width           =   2175
         End
         Begin VB.PictureBox Command3D6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5160
            ScaleHeight     =   465
            ScaleWidth      =   1065
            TabIndex        =   1
            Top             =   960
            Width           =   1095
         End
         Begin VB.PictureBox Command3D5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   1785
            TabIndex        =   2
            Top             =   1560
            Width           =   1815
         End
         Begin VB.PictureBox Command3D4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   1785
            TabIndex        =   3
            Top             =   960
            Width           =   1815
         End
         Begin VB.PictureBox Command3D3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2520
            ScaleHeight     =   465
            ScaleWidth      =   2145
            TabIndex        =   9
            Top             =   360
            Width           =   2175
         End
         Begin VB.PictureBox Command3D2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2520
            ScaleHeight     =   465
            ScaleWidth      =   2145
            TabIndex        =   10
            Top             =   960
            Width           =   2175
         End
         Begin VB.PictureBox Command3d1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   240
            ScaleHeight     =   465
            ScaleWidth      =   1785
            TabIndex        =   11
            Top             =   360
            Width           =   1815
         End
         Begin VB.PictureBox Command3D9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   5160
            ScaleHeight     =   465
            ScaleWidth      =   2385
            TabIndex        =   19
            Top             =   360
            Width           =   2415
         End
      End
      Begin VB.PictureBox Panel3D1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1800
         ScaleHeight     =   480
         ScaleWidth      =   4545
         TabIndex        =   16
         Top             =   1920
         Width           =   4575
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   55
            MaxLength       =   20
            TabIndex        =   17
            Top             =   60
            Width           =   4455
         End
      End
      Begin VB.PictureBox Panel3D2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   1800
         ScaleHeight     =   480
         ScaleWidth      =   1305
         TabIndex        =   14
         Top             =   720
         Width           =   1335
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   60
            MaxLength       =   4
            TabIndex        =   15
            Top             =   60
            Width           =   1215
         End
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Leyenda"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1800
         TabIndex        =   12
         Top             =   1560
         Width           =   870
      End
   End
   Begin VB.PictureBox Panel3D3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      ScaleHeight     =   720
      ScaleWidth      =   9225
      TabIndex        =   4
      Top             =   240
      Width           =   9255
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2160
         TabIndex        =   7
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
   End
End
Attribute VB_Name = "LEYENDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3D1_Click()
 If DN_LEYENDAS.RecordCount > 0 Then
  If Not DN_LEYENDAS.EOF Then
    DN_LEYENDAS.MoveNext
    If Not DN_LEYENDAS.EOF Then
      VER_LEYENDAS
    Else
      DN_LEYENDAS.MoveLast
    End If
  Else
    DN_LEYENDAS.MoveLast
  End If
 End If
End Sub

Private Sub Command3D1_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub Command3D10_Click()
 Target.Show 1
 
 If (DESTINO_REPORTE = REPORTE_PANTALLA) Or (DESTINO_REPORTE = REPORTE_IMPRESORA) Then
  listados.Report1.DataFiles(0) = PATH_PROG + "EMPR" + MAIN.Label2.Caption + ".MDB"
  listados.Report1.ReportFileName = PATH_PROG + "LEYENDAS.RPT"
  listados.Report1.SelectionFormula = "{" + TABLA_LEYENDAS + "." + CODIGO_LEYENDA + "} > ''"
  listados.Report1.WindowTitle = "Listado de Leyendas"
  listados.Report1.WindowTop = -3
  listados.Report1.WindowLeft = -3
  listados.Report1.WindowWidth = Screen.Width
  listados.Report1.WindowHeight = Screen.Height

  listados.Report1.SortFields(0) = "+{" + TABLA_LEYENDAS + "." + CODIGO_LEYENDA + "}"

  listados.Report1.Destination = DESTINO_REPORTE
  listados.Report1.Action = 1
 End If

End Sub

Private Sub Command3D2_Click()
 If DN_LEYENDAS.RecordCount > 0 Then
   DN_LEYENDAS.MoveLast
   VER_LEYENDAS
 End If
End Sub

Private Sub Command3D2_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub Command3D3_Click()
 If DN_LEYENDAS.RecordCount > 0 Then
   DN_LEYENDAS.MoveFirst
   VER_LEYENDAS
 End If
End Sub

Private Sub Command3D3_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub Command3D4_Click()
 If DN_LEYENDAS.RecordCount > 0 Then
  If Not DN_LEYENDAS.BOF Then
    DN_LEYENDAS.MovePrevious
    If Not DN_LEYENDAS.BOF Then
      VER_LEYENDAS
    Else
      DN_LEYENDAS.MoveFirst
    End If
  Else
    DN_LEYENDAS.MoveFirst
  End If
 End If
End Sub

Private Sub Command3D4_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub COMMAND3D5_CLICK()
  COD = UCase$(LTrim$(RTrim$(LEYENDAS.Text1.text)))
  NOM = UCase$(LTrim$(RTrim$(LEYENDAS.Text2.text)))
  
  If COD = "" Or NOM = "" Then
    Beep
    MsgBox "Los Datos están Incompletos.", 48, "Error"
    Exit Sub
  End If

  CAD = CODIGO_LEYENDA + " = '" + COD + "'"
  DN_LEYENDAS.FindFirst CAD
  
  If DN_LEYENDAS.NoMatch Then
    'ALTA
    DN_LEYENDAS.AddNew
    Text1.text = ""
    Text2.text = ""
  Else
    'MODIFICACIÓN
    DN_LEYENDAS.Edit
  End If

  DN_LEYENDAS(CODIGO_LEYENDA) = COD
  DN_LEYENDAS(LEYENDA) = NOM
  DN_LEYENDAS(LEYENDA_RPT_EMPRESA) = MAIN.Label3.Caption
  DN_LEYENDAS.Update
  COMMAND3D9_CLICK
End Sub

Private Sub Command3D5_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub COMMAND3D6_CLICK()
  listados.Caption = "Lista de Leyendas"
  VER_LISTA_LEYENDAS
End Sub

Private Sub Command3D6_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub COMMAND3D7_CLICK()
 Text1.text = ""
 Text2.text = ""
 LEYENDAS.Hide
 DN_LEYENDAS.Close
 DB_LEYENDAS.Close
 MAIN.Refresh
End Sub

Private Sub Command3D7_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub COMMAND3D8_CLICK()
 CAD = UCase$(LTrim$(RTrim$(LEYENDAS.Text1.text)))
 CAD = CODIGO_LEYENDA + " = '" + CAD + "'"
 DN_LEYENDAS.FindFirst CAD

 LEYE = DN_LEYENDAS(LEYENDA) & ""
 A = MsgBox("Desea borrar la Leyenda " + LEYE + " ?", 32 + 4, "Borrar Leyenda")
 If A = 7 Then Exit Sub

 If DN_LEYENDAS.NoMatch Then
   Beep
   MsgBox "La leyenda no pudo ser borrada.", 48, "Error"
 Else
   DN_LEYENDAS.Delete
   COMMAND3D9_CLICK
 End If
End Sub

Private Sub Command3D8_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub COMMAND3D9_CLICK()
  Text1.text = ""
  Text2.text = ""
  Text1.SetFocus
End Sub

Private Sub Form_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK

End Sub

Private Sub TEXT1_KeyPress(KEYASCII As Integer)
  If KEYASCII = 13 Then
    Text2.text = ""
    C = LTrim$(RTrim$(Text1.text))
    CAD = CODIGO_LEYENDA + " = '" + C + "'"
    DN_LEYENDAS.FindFirst CAD
    If Not DN_LEYENDAS.NoMatch Then
       Text2.text = DN_LEYENDAS(LEYENDA) & ""
    End If
    Text2.SetFocus
  End If
  If KEYASCII = 27 Then COMMAND3D7_CLICK
End Sub

Private Sub Text2_KeyDown(KEYCODE As Integer, Shift As Integer)
  If KEYCODE = KEY_F4 Then
    Text2.text = ""
    Text1.SetFocus
  End If
End Sub

Private Sub Text2_KeyPress(KEYASCII As Integer)
  If KEYASCII = 27 Then COMMAND3D7_CLICK
  If KEYASCII = 13 Then Command3D5.SetFocus
End Sub

