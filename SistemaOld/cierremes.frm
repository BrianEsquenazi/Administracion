VERSION 5.00
Begin VB.Form PrgCierre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre del Mes"
   ClientHeight    =   2415
   ClientLeft      =   3105
   ClientTop       =   990
   ClientWidth     =   5805
   FillColor       =   &H00800000&
   Icon            =   "cierremes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1426.862
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   5450.582
   ShowInTaskbar   =   0   'False
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
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   2895
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
      Left            =   2400
      MaxLength       =   4
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Mes 
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
      Left            =   1800
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   360
      Width           =   495
   End
   Begin VB.Image Cancela 
      Height          =   480
      Left            =   3000
      MouseIcon       =   "cierremes.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "cierremes.frx":074C
      ToolTipText     =   "Menu Principal"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image Acepta 
      Height          =   480
      Left            =   1920
      MouseIcon       =   "cierremes.frx":0F8E
      MousePointer    =   99  'Custom
      Picture         =   "cierremes.frx":1298
      ToolTipText     =   "Confirma el Proceso"
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Mes / Año"
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
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "PrgCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Acepta_Click()

    ZSql = "Select *"
    ZSql = ZSql + " FROM Cierre"
    ZSql = ZSql + " Where Cierre.Mes = " + "'" + Mes.Text + "'"
    ZSql = ZSql + " and Cierre.Ano = " + "'" + Ano.Text + "'"
    spCierre = ZSql
    Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
    If rstCierre.RecordCount > 0 Then
    
        rstCierre.Close
        
        ZSql = "UPDATE Cierre SET "
        ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "'"
        ZSql = ZSql + " Where Mes = " + "'" + Mes.Text + "'"
        ZSql = ZSql + " and Ano = " + "'" + Ano.Text + "'"
        spCierre = ZSql
        Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
        
            Else
            
        ZSql = "INSERT INTO Cierre ("
        ZSql = ZSql + "Mes ,"
        ZSql = ZSql + "Ano ,"
        ZSql = ZSql + "Estado)"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + Mes.Text + "',"
        ZSql = ZSql + "'" + Ano.Text + "',"
        ZSql = ZSql + "'" + Str$(Estado.ListIndex) + "')"
        spCierre = ZSql
        Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    Call Cancela_Click

End Sub

Private Sub Cancela_Click()
    PrgCierre.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()
        
    Estado.Clear
    
    Estado.AddItem "Mes Abierto"
    Estado.AddItem "Mes Cerrado"
    Estado.AddItem ""
    
    Estado.ListIndex = 2

    Mes.Text = ""
    Ano.Text = ""
    
End Sub

Private Sub Mes_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = "Select *"
        ZSql = ZSql + " FROM Cierre"
        ZSql = ZSql + " Where Cierre.Mes = " + "'" + Mes.Text + "'"
        ZSql = ZSql + " and Cierre.Ano = " + "'" + Ano.Text + "'"
        spCierre = ZSql
        Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
        If rstCierre.RecordCount > 0 Then
            Estado.ListIndex = rstCierre!Estado
            rstCierre.Close
        End If
    
        Ano.SetFocus
    End If
    If KeyAscii = 27 Then
        Mes.Text = ""
    End If
End Sub

Private Sub Ano_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        ZSql = "Select *"
        ZSql = ZSql + " FROM Cierre"
        ZSql = ZSql + " Where Cierre.Mes = " + "'" + Mes.Text + "'"
        ZSql = ZSql + " and Cierre.Ano = " + "'" + Ano.Text + "'"
        spCierre = ZSql
        Set rstCierre = db.OpenRecordset(spCierre, dbOpenSnapshot, dbSQLPassThrough)
        If rstCierre.RecordCount > 0 Then
            Estado.ListIndex = rstCierre!Estado
            rstCierre.Close
        End If
    
        Mes.SetFocus
    End If
    If KeyAscii = 27 Then
        Ano.Text = ""
    End If
End Sub


