VERSION 5.00
Begin VB.Form PrgCheckList 
   Caption         =   "Check-List del Transportista"
   ClientHeight    =   8595
   ClientLeft      =   2025
   ClientTop       =   480
   ClientWidth     =   8715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
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
      TabIndex        =   3
      Top             =   120
      Width           =   8295
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
         Left            =   2880
         TabIndex        =   48
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox DesExpreso 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4920
         MaxLength       =   50
         TabIndex        =   0
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
         Caption         =   "Confirma"
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
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   5
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
         TabIndex        =   4
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
Attribute VB_Name = "PrgCheckList"
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

Dim ZVector(100) As String

Private Sub Confirma_Click()

    ZLugar = 0
    Erase ZVector
    ZPeligrosa = "N"
    
    spInforme = "ListaInforme " + "'" + WPasaInforme + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If .EOF = False Then
                
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar) = rstInforme!Articulo
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZArticulo = ZVector(Ciclo)
        ZClase = ""
                
        spArticulo = "ConsultaArticulo " + "'" + ZArticulo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            rstArticulo.Close
        End If
        
        If Trim(ZClase) <> "" Then
            ZPeligrosa = "S"
        End If
        
    Next Ciclo

    Suma1 = 0
    Suma2 = 0
    Suma3 = 0
    Suma4 = 0
    Suma5 = 0
    Suma6 = 0
    Suma7 = 0
    Suma8 = 0
    
    ZZItem1 = ""
    ZZItem2 = ""
    ZZItem3 = ""
    ZZItem4 = ""
    ZZItem5 = ""
    ZZItem6 = ""
    ZZItem7 = ""
    ZZItem8 = ""

    Suma1 = Item11.Value + Item12.Value + Item13.Value
    Suma2 = Item21.Value + Item22.Value + Item23.Value
    Suma3 = Item31.Value + Item32.Value + Item33.Value
    Suma4 = Item41.Value + Item42.Value + Item43.Value
    Suma5 = Item51.Value + Item52.Value + Item53.Value
    Suma6 = Item61.Value + Item62.Value + Item63.Value
    Suma7 = Item71.Value + Item72.Value + Item73.Value
    Suma8 = Item81.Value + Item82.Value + Item83.Value
    
    If Suma1 <> 1 Or Suma2 <> 1 Or Suma3 <> 1 Or Suma4 <> 1 Or Suma5 <> 1 Or Suma6 <> 1 Or Suma7 <> 1 Or Suma8 <> 1 Then
        m$ = "Error en la carga de datos"
        G% = MsgBox(m$, 0, "Check-List del Transportista")
        Exit Sub
    End If
    
    If Expreso.ListIndex <= 0 Then
        m$ = "Error en la carga de datos no hay Expreso"
        G% = MsgBox(m$, 0, "Check-List del Transportista")
        Exit Sub
    End If
    
    If ZPeligrosa = "S" Then
        If Item63.Value = 1 Or Item73.Value = 1 Or Item83.Value = 1 Then
            m$ = "Error en la carga de datos, al incluir carga peligrosa " + Chr$(13) + "es obligatorio la carga de todos los datos"
            G% = MsgBox(m$, 0, "Check-List del Transportista")
            Exit Sub
        End If
        If Trim(Placa.Text) = "" Then
            m$ = "Error en la carga de datos, se debe informar la Placa"
            G% = MsgBox(m$, 0, "Check-List del Transportista")
            Exit Sub
        End If
        If Trim(Rombo.Text) = "" Then
            m$ = "Error en la carga de datos, se debe informar el Rombo"
            G% = MsgBox(m$, 0, "Check-List del Transportista")
            Exit Sub
        End If
    End If
    
    If Item11.Value = 1 Then
        ZZItem1 = "1"
    End If
    If Item12.Value = 1 Then
        ZZItem1 = "2"
    End If
    If Item13.Value = 1 Then
        ZZItem1 = "3"
    End If
    
    If Item21.Value = 1 Then
        ZZItem2 = "1"
    End If
    If Item22.Value = 1 Then
        ZZItem2 = "2"
    End If
    If Item23.Value = 1 Then
        ZZItem2 = "3"
    End If
    
    If Item31.Value = 1 Then
        ZZItem3 = "1"
    End If
    If Item32.Value = 1 Then
        ZZItem3 = "2"
    End If
    If Item33.Value = 1 Then
        ZZItem3 = "3"
    End If
    
    If Item41.Value = 1 Then
        ZZItem4 = "1"
    End If
    If Item42.Value = 1 Then
        ZZItem4 = "2"
    End If
    If Item43.Value = 1 Then
        ZZItem4 = "3"
    End If
    
    If Item51.Value = 1 Then
        ZZItem5 = "1"
    End If
    If Item52.Value = 1 Then
        ZZItem5 = "2"
    End If
    If Item53.Value = 1 Then
        ZZItem5 = "3"
    End If
    
    If Item61.Value = 1 Then
        ZZItem6 = "1"
    End If
    If Item62.Value = 1 Then
        ZZItem6 = "2"
    End If
    If Item63.Value = 1 Then
        ZZItem6 = "3"
    End If
    
    If Item71.Value = 1 Then
        ZZItem7 = "1"
    End If
    If Item72.Value = 1 Then
        ZZItem7 = "2"
    End If
    If Item73.Value = 1 Then
        ZZItem7 = "3"
    End If
    
    If Item81.Value = 1 Then
        ZZItem8 = "1"
    End If
    If Item82.Value = 1 Then
        ZZItem8 = "2"
    End If
    If Item83.Value = 1 Then
        ZZItem8 = "3"
    End If


    ZSql = ""
    ZSql = ZSql + "UPDATE Informe SET "
    ZSql = ZSql + " Expreso = " + "'" + Str$(Expreso.ListIndex) + "',"
    ZSql = ZSql + " DesExpreso = " + "'" + DesExpreso.Text + "',"
    ZSql = ZSql + " Chapa = " + "'" + Chapa.Text + "',"
    ZSql = ZSql + " Chofer = " + "'" + Chofer.Text + "',"
    ZSql = ZSql + " Item1 = " + "'" + ZZItem1 + "',"
    ZSql = ZSql + " Item2 = " + "'" + ZZItem2 + "',"
    ZSql = ZSql + " Item3 = " + "'" + ZZItem3 + "',"
    ZSql = ZSql + " Item4 = " + "'" + ZZItem4 + "',"
    ZSql = ZSql + " Item5 = " + "'" + ZZItem5 + "',"
    ZSql = ZSql + " Item6 = " + "'" + ZZItem6 + "',"
    ZSql = ZSql + " Item7 = " + "'" + ZZItem7 + "',"
    ZSql = ZSql + " Item8 = " + "'" + ZZItem8 + "',"
    ZSql = ZSql + " Placa = " + "'" + Placa.Text + "',"
    ZSql = ZSql + " Rombo = " + "'" + Rombo.Text + "',"
    ZSql = ZSql + " Observaciones = " + "'" + Observaciones.Text + "'"
    ZSql = ZSql + " Where Informe = " + "'" + WPasaInforme + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    ZZPaso = ZZPasaProceso

    ZZPasaProceso = 1
    PrgCheckList.Hide
    Unload Me
    If ZZPaso = 1 Then
        PrgInforme.Show
            Else
        PrgActualizaInforme.Show
    End If

End Sub



Private Sub Expreso_Click()
Rem BY NAN
ver = Expreso.ListIndex

If ver = "3" Then
DesExpreso.SetFocus
 DesExpreso.BackColor = &HFFFF80
 Rem BY NAN
End If
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

End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Private Sub DesExpreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Chapa.SetFocus
    End If
    If KeyAscii = 27 Then
        DesExpreso.Text = ""
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

