VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form PrgConfigInversion 
   AutoRedraw      =   -1  'True
   Caption         =   "Ingreos de Atributos de Desarrollo"
   ClientHeight    =   8400
   ClientLeft      =   225
   ClientTop       =   390
   ClientWidth     =   11535
   LinkTopic       =   "Form2"
   ScaleHeight     =   8400
   ScaleWidth      =   11535
   Begin VB.TextBox Operador 
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Salida 
      Caption         =   "Salida"
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
      Left            =   6360
      TabIndex        =   3
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
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
      Left            =   4080
      TabIndex        =   2
      Top             =   7080
      Width           =   1215
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   6135
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   10821
      _Version        =   327680
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "ConfigInversion.frx":0000
      Tab(0).ControlCount=   10
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Titulo1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Titulo2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Titulo4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Titulo3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Titulo5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Opcion1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Opcion2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Opcion4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Opcion3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Opcion5"
      Tab(0).Control(9).Enabled=   0   'False
      Begin VB.CheckBox Opcion5 
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
         Left            =   4440
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.CheckBox Opcion3 
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
         Left            =   4440
         TabIndex        =   11
         Top             =   1140
         Width           =   615
      End
      Begin VB.CheckBox Opcion4 
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
         Left            =   4440
         TabIndex        =   10
         Top             =   1500
         Width           =   615
      End
      Begin VB.CheckBox Opcion2 
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
         Left            =   4440
         TabIndex        =   9
         Top             =   780
         Width           =   615
      End
      Begin VB.CheckBox Opcion1 
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
         Left            =   4440
         TabIndex        =   7
         Top             =   420
         Width           =   615
      End
      Begin VB.Label Titulo5 
         Caption         =   "Ingreso de "
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   3975
      End
      Begin VB.Label Titulo3 
         Caption         =   "Ingrso de "
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
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Titulo4 
         Caption         =   "Ingreso de "
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
         Left            =   240
         TabIndex        =   12
         Top             =   1500
         Width           =   3975
      End
      Begin VB.Label Titulo2 
         Caption         =   "Ingreso de "
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
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   3975
      End
      Begin VB.Label Titulo1 
         Caption         =   "Ingrso de "
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
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Operador"
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
      TabIndex        =   4
      Top             =   180
      Width           =   2295
   End
End
Attribute VB_Name = "PrgConfigInversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstOperador As Recordset
Dim spOperador As String
Dim rstAtributos As Recordset
Dim spAtributos As String
Dim XParam As String


Sub Form_Load()

    Operador.Text = ""
    DesOperador.Caption = ""
    
    Tablas.TabCaption(0) = "Maestros"
    
    Rem titulo1
    
    Titulo1.Caption = "Ingreso de Sectores"
    Titulo2.Caption = "Ingreso de Proyectos"
    Titulo3.Caption = "Asignacion de Proyectos al Año"
    Titulo4.Caption = "Ingreso de Avance de los Proyectos"
    Titulo5.Caption = "Listado de Avance de Proyecto"
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0

End Sub


Private Sub Graba_Click()

    XParam = "'" + Operador.Text + "','" _
                 + "5" + "'"
    spAtributos = "BorrarAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    WAtributo1 = ""
    WAtributo2 = ""
    WAtributo3 = ""
    WAtributo4 = ""
    WAtributo5 = ""
    WAtributo6 = ""
    WAtributo7 = ""
    WAtributo8 = ""
    WAtributo9 = ""
    WAtributo10 = ""
    
    
    
    If Opcion1.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion2.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion3.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion4.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    If Opcion5.Value = 0 Then
        WAtributo1 = WAtributo1 + "0"
            Else
        WAtributo1 = WAtributo1 + "1"
    End If
    
    WProceso = "5"
                                       
    XParam = "'" + Operador.Text + "','" _
                 + WProceso + "','" _
                 + WAtributo1 + "','" _
                 + WAtributo2 + "','" _
                 + WAtributo3 + "','" _
                 + WAtributo4 + "','" _
                 + WAtributo5 + "','" _
                 + WAtributo6 + "','" _
                 + WAtributo7 + "','" _
                 + WAtributo8 + "','" _
                 + WAtributo9 + "','" _
                 + WAtributo10 + "'"
                    
    spAtributos = "AltaAtributos " + XParam
    Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
    
    Operador.Text = ""
    DesOperador.Caption = ""
    
    Opcion1.Value = 0
    Opcion2.Value = 0
    Opcion3.Value = 0
    Opcion4.Value = 0
    Opcion5.Value = 0
    
    Operador.SetFocus
    
    Tablas.Tab = 0

End Sub

Private Sub Operador_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Operador.Text <> "" Then
        
            spOperador = "ConsultaOperador " + "'" + Operador.Text + "'"
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                DesOperador.Caption = rstOperador!Descripcion
                rstOperador.Close
                
                Opcion1.Value = 0
                Opcion2.Value = 0
                Opcion3.Value = 0
                Opcion4.Value = 0
                Opcion5.Value = 0
                
                XParam = "'" + Operador.Text + "','" _
                             + "5" + "'"
                spAtributos = "ConsultaAtributo " + XParam
                Set rstAtributos = db.OpenRecordset(spAtributos, dbOpenSnapshot, dbSQLPassThrough)
                If rstAtributos.RecordCount > 0 Then
                    Opcion1.Value = Val(Mid$(rstAtributos!atributo1, 1, 1))
                    Opcion2.Value = Val(Mid$(rstAtributos!atributo1, 2, 1))
                    Opcion3.Value = Val(Mid$(rstAtributos!atributo1, 3, 1))
                    Opcion4.Value = Val(Mid$(rstAtributos!atributo1, 4, 1))
                    Opcion5.Value = Val(Mid$(rstAtributos!atributo1, 5, 1))
                    rstAtributos.Close
                End If
                
            End If
            
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Salida_Click()
    PrgConfigInversion.Hide
    Unload Me
    Menu.Show
End Sub

