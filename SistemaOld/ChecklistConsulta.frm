VERSION 5.00
Begin VB.Form PrgCheckListConsulta 
   Caption         =   "CONSULTA DE Check-List del Transportista"
   ClientHeight    =   8595
   ClientLeft      =   2025
   ClientTop       =   480
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   8715
   Begin VB.Frame PantaEnvase 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.TextBox DesExpreso 
         Height          =   285
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   48
         Top             =   360
         Width           =   3015
      End
      Begin VB.CheckBox Item83 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   47
         Top             =   6240
         Width           =   735
      End
      Begin VB.CheckBox Item82 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   46
         Top             =   6240
         Width           =   615
      End
      Begin VB.CheckBox Item81 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   45
         Top             =   6240
         Width           =   495
      End
      Begin VB.CheckBox Item73 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   41
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox Item72 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   39
         Top             =   5040
         Width           =   615
      End
      Begin VB.CheckBox Item71 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   38
         Top             =   5040
         Width           =   495
      End
      Begin VB.TextBox Rombo 
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
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   37
         Text            =   " "
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox Placa 
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
         Left            =   6000
         MaxLength       =   20
         TabIndex        =   36
         Text            =   " "
         Top             =   5520
         Width           =   1815
      End
      Begin VB.CheckBox Item63 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   32
         Top             =   4560
         Width           =   735
      End
      Begin VB.CheckBox Item62 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   30
         Top             =   4560
         Width           =   615
      End
      Begin VB.CheckBox Item61 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   29
         Top             =   4560
         Width           =   495
      End
      Begin VB.CheckBox Item53 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   28
         Top             =   3480
         Width           =   735
      End
      Begin VB.CheckBox Item43 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   27
         Top             =   3000
         Width           =   735
      End
      Begin VB.CheckBox Item33 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   26
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox Observaciones 
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
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   6720
         Width           =   3015
      End
      Begin VB.CheckBox Item52 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   16
         Top             =   3480
         Width           =   615
      End
      Begin VB.CheckBox Item51 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   15
         Top             =   3480
         Width           =   495
      End
      Begin VB.CheckBox Item42 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   14
         Top             =   3000
         Width           =   615
      End
      Begin VB.CheckBox Item41 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   13
         Top             =   3000
         Width           =   495
      End
      Begin VB.CheckBox Item32 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   12
         Top             =   2520
         Width           =   615
      End
      Begin VB.CheckBox Item31 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   11
         Top             =   2520
         Width           =   495
      End
      Begin VB.CheckBox Item12 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
      Begin VB.CheckBox Item11 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   9
         Top             =   1560
         Width           =   495
      End
      Begin VB.CheckBox Item22 
         Caption         =   "No"
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
         Left            =   5880
         MaskColor       =   &H00FF0000&
         TabIndex        =   8
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox Item21 
         Caption         =   "Si"
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
         Left            =   4920
         MaskColor       =   &H00FF0000&
         TabIndex        =   7
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Confirma 
         Caption         =   "Cierra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4080
         TabIndex        =   6
         Top             =   7320
         Width           =   1335
      End
      Begin VB.ComboBox Expreso 
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
         Left            =   2760
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.TextBox Chapa 
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
         Left            =   4920
         MaxLength       =   20
         TabIndex        =   4
         Text            =   " "
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Chofer 
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
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   3
         Text            =   " "
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CheckBox Item13 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   2
         Top             =   1560
         Width           =   735
      End
      Begin VB.CheckBox Item23 
         Caption         =   "N/a"
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
         Left            =   6960
         MaskColor       =   &H00FF0000&
         TabIndex        =   1
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "REQUISITOS A CUMPLIR SOLO PARA EL TRANSPORTE DE SUSTANCIAS PELIGROSAS"
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
         Height          =   285
         Left            =   360
         TabIndex        =   44
         Top             =   3960
         Width           =   7815
      End
      Begin VB.Label Label10 
         Caption         =   "Hojas de Seguridad en Castellano de cada Producto Transportado"
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
         Height          =   435
         Left            =   360
         TabIndex        =   43
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "Observaciones"
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
         Height          =   405
         Left            =   360
         TabIndex        =   42
         Top             =   6840
         Width           =   4335
      End
      Begin VB.Label Label8 
         Caption         =   "Guia de Intervencion d cada Sustancia Peligrosa Transportada"
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
         Height          =   525
         Left            =   360
         TabIndex        =   40
         Top             =   6240
         Width           =   4335
      End
      Begin VB.Label Label7 
         Caption         =   "Rombo"
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
         Left            =   4920
         TabIndex        =   35
         Top             =   5880
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Placa"
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
         Left            =   4920
         TabIndex        =   34
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Placas y Rombos segun Nro. ONU de las sustancias peligrosas Transportads"
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
         Height          =   525
         Left            =   360
         TabIndex        =   33
         Top             =   5520
         Width           =   4335
      End
      Begin VB.Label Label4 
         Caption         =   "Habilitacion del Camion para Transporte de Sustancias Peligrosas"
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
         Height          =   525
         Left            =   360
         TabIndex        =   31
         Top             =   4920
         Width           =   4335
      End
      Begin VB.Label DescriEnsayo5 
         Caption         =   "Licencia de Conductor para Transporte de Sustancias Peligrosas"
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
         Height          =   525
         Left            =   360
         TabIndex        =   25
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label DescriEnsayo4 
         Caption         =   "Matafuegos de Kilos para la caja"
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
         Height          =   285
         Left            =   360
         TabIndex        =   24
         Top             =   3000
         Width           =   4455
      End
      Begin VB.Label DescriEnsayo3 
         Caption         =   "Simiremolque Adecuado al Largo del Contenedor"
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
         Height          =   435
         Left            =   360
         TabIndex        =   23
         Top             =   2040
         Width           =   4335
      End
      Begin VB.Label DescriEnsayo1 
         Caption         =   "Correcto Anclaje (4 lados) del Contenedor en el Semiremolque"
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
         Height          =   435
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label DescriEnsayo2 
         Caption         =   "Elementos para Contencion de Derrames"
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
         Height          =   435
         Left            =   360
         TabIndex        =   21
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor de Transporte"
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
         Height          =   315
         Left            =   360
         TabIndex        =   20
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label2 
         Caption         =   "Chapa Patente de la Unidad"
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
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre y Apellido del Conductor"
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
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   4455
      End
   End
End
Attribute VB_Name = "PrgCheckListConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZZItem1 As String
Dim ZZItem2 As String
Dim ZZItem3 As String
Dim ZZItem4 As String
Dim ZZItem5 As String
Dim ZZItem6 As String
Dim ZZItem7 As String
Dim ZZItem8 As String

Private Sub Confirma_Click()
    ZZPasaProceso = 99
    PrgCheckListConsulta.Hide
    Unload Me
    PrgInforme.Show
End Sub

Private Sub Form_Load()

    Expreso.Clear
    
    Expreso.AddItem ""
    Expreso.AddItem "FURTADO"
    Expreso.AddItem "BERN"
    Expreso.AddItem "OTRO"
    
    Expreso.ListIndex = 0

    Item11.Value = 0
    Item12.Value = 0
    Item13.Value = 0
    Item21.Value = 0
    Item22.Value = 0
    Item23.Value = 0
    Item31.Value = 0
    Item32.Value = 0
    Item33.Value = 0
    Item41.Value = 0
    Item42.Value = 0
    Item43.Value = 0
    Item51.Value = 0
    Item52.Value = 0
    Item53.Value = 0
    Item61.Value = 0
    Item62.Value = 0
    Item63.Value = 0
    Item71.Value = 0
    Item72.Value = 0
    Item73.Value = 0
    Item81.Value = 0
    Item82.Value = 0
    Item83.Value = 0
    
    Chapa.Text = ""
    Chofer.Text = ""
    Placa.Text = ""
    Rombo.Text = ""
    Observaciones.Text = ""
    DesExpreso.Text = ""
    
    spInforme = "ListaInforme " + "'" + WPasaInforme + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        Expreso.ListIndex = IIf(IsNull(rstInforme!Expreso), "0", rstInforme!Expreso)
        Chapa.Text = IIf(IsNull(rstInforme!Chapa), "", rstInforme!Chapa)
        Chofer.Text = IIf(IsNull(rstInforme!Chofer), "", rstInforme!Chofer)
        Placa.Text = IIf(IsNull(rstInforme!Placa), "", rstInforme!Placa)
        Rombo.Text = IIf(IsNull(rstInforme!Rombo), "", rstInforme!Rombo)
        Observaciones.Text = IIf(IsNull(rstInforme!Observaciones), "", rstInforme!Observaciones)
        DesExpreso.Text = IIf(IsNull(rstInforme!DesExpreso), "", rstInforme!DesExpreso)
        ZZItem1 = IIf(IsNull(rstInforme!Item1), "0", rstInforme!Item1)
        ZZItem2 = IIf(IsNull(rstInforme!Item2), "0", rstInforme!Item2)
        ZZItem3 = IIf(IsNull(rstInforme!Item3), "0", rstInforme!Item3)
        ZZItem4 = IIf(IsNull(rstInforme!Item4), "0", rstInforme!Item4)
        ZZItem5 = IIf(IsNull(rstInforme!Item5), "0", rstInforme!Item5)
        ZZItem6 = IIf(IsNull(rstInforme!Item6), "0", rstInforme!Item6)
        ZZItem7 = IIf(IsNull(rstInforme!Item7), "0", rstInforme!Item7)
        ZZItem8 = IIf(IsNull(rstInforme!Item8), "0", rstInforme!Item8)
        rstInforme.Close
    End If
    
    If Val(ZZItem1) = 1 Then
        Item11.Value = 1
    End If
    If Val(ZZItem1) = 2 Then
        Item12.Value = 1
    End If
    If Val(ZZItem1) = 3 Then
        Item13.Value = 1
    End If
    
    If Val(ZZItem2) = 1 Then
        Item21.Value = 1
    End If
    If Val(ZZItem2) = 2 Then
        Item22.Value = 1
    End If
    If Val(ZZItem2) = 3 Then
        Item23.Value = 1
    End If
    
    If Val(ZZItem3) = 1 Then
        Item31.Value = 1
    End If
    If Val(ZZItem3) = 2 Then
        Item32.Value = 1
    End If
    If Val(ZZItem3) = 3 Then
        Item33.Value = 1
    End If
    
    If Val(ZZItem4) = 1 Then
        Item41.Value = 1
    End If
    If Val(ZZItem4) = 2 Then
        Item42.Value = 1
    End If
    If Val(ZZItem4) = 3 Then
        Item43.Value = 1
    End If
    
    If Val(ZZItem5) = 1 Then
        Item51.Value = 1
    End If
    If Val(ZZItem5) = 2 Then
        Item52.Value = 1
    End If
    If Val(ZZItem5) = 3 Then
        Item53.Value = 1
    End If
    
    If Val(ZZItem6) = 1 Then
        Item61.Value = 1
    End If
    If Val(ZZItem6) = 2 Then
        Item62.Value = 1
    End If
    If Val(ZZItem6) = 3 Then
        Item63.Value = 1
    End If
    
    If Val(ZZItem7) = 1 Then
        Item71.Value = 1
    End If
    If Val(ZZItem7) = 2 Then
        Item72.Value = 1
    End If
    If Val(ZZItem7) = 3 Then
        Item73.Value = 1
    End If
    
    If Val(ZZItem8) = 1 Then
        Item81.Value = 1
    End If
    If Val(ZZItem8) = 2 Then
        Item82.Value = 1
    End If
    If Val(ZZItem8) = 3 Then
        Item83.Value = 1
    End If
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


Private Sub Chapa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Chofer.SetFocus
    End If
    If KeyAscii = 27 Then
        Chapa.Text = ""
    End If
End Sub

Private Sub Chofer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Placa.SetFocus
    End If
    If KeyAscii = 27 Then
        Chofer.Text = ""
    End If
End Sub

Private Sub Placa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Rombo.SetFocus
    End If
    If KeyAscii = 27 Then
        Placa.Text = ""
    End If
End Sub

Private Sub Rombo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
    If KeyAscii = 27 Then
        Rombo.Text = ""
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Chapa.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

