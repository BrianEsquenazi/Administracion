VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConsultaHojaTotal 
   AutoRedraw      =   -1  'True
   Caption         =   "Supervision de Produccion"
   ClientHeight    =   8325
   ClientLeft      =   105
   ClientTop       =   345
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   11850
   Begin VB.Timer Timer1 
      Left            =   8400
      Top             =   5640
   End
   Begin VB.Frame Box8 
      Height          =   2415
      Left            =   4080
      TabIndex        =   65
      Top             =   5160
      Width           =   3615
      Begin VB.Image Foto8 
         Height          =   855
         Left            =   360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label ImpreHora8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   73
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   72
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreOperador8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   71
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   70
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreProducto8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   69
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreHoja8 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   68
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreTemperatura8 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   67
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 8"
         Height          =   255
         Left            =   360
         TabIndex        =   66
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Box7 
      Height          =   2415
      Left            =   120
      TabIndex        =   56
      Top             =   5160
      Width           =   3615
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 7"
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   63
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   62
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   61
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   60
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   59
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   58
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora7 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   57
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto7 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box4 
      Height          =   2415
      Left            =   120
      TabIndex        =   47
      Top             =   2640
      Width           =   3615
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 4"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura4 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   54
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   53
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   52
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   51
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   50
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora4 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   48
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto4 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box5 
      Height          =   2415
      Left            =   4080
      TabIndex        =   38
      Top             =   2640
      Width           =   3615
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 5"
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   45
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   43
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   42
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   41
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora5 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto5 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box6 
      Height          =   2415
      Left            =   7920
      TabIndex        =   29
      Top             =   2640
      Width           =   3615
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 6"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura6 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   36
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   35
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   34
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   33
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   32
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   31
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora6 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   30
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto6 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box3 
      Height          =   2415
      Left            =   7920
      TabIndex        =   20
      Top             =   120
      Width           =   3615
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 3"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura3 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   27
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   26
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   25
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   24
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   23
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   22
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora3 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   21
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto3 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box2 
      Height          =   2415
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   3615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 2"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura2 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   17
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   16
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   13
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora2 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto2 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box1 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 1"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreOperador1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreTemperatura1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHora1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreProducto1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreHoja1 
         BackColor       =   &H00FFFF80&
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
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Image Foto1 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame CambiaOperador 
      Height          =   975
      Left            =   4440
      TabIndex        =   0
      Top             =   3120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox ListaCambiaOperador 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4155
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3975
      End
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7800
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5760
      MouseIcon       =   "consultahojatotal.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "consultahojatotal.frx":030A
      ToolTipText     =   "Salida"
      Top             =   7680
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaHojaTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim XParam As String
Dim WGraba As String
Dim XEmpresa As String
Dim ZDatos(100, 30) As String
Dim ZOperador(1000, 2) As String
Dim ZZProceso As String
Dim ZZOperario As String
Dim ZZCambiaOperario As Integer

Private Sub CancelaGraba_Click()
    Call cmdClose_Click
End Sub

Private Sub cmdClose_Click()
    PrgConsultaHojaTotal.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Operario_KeyPress(KeyAscii As Integer)
    Call Proceso_Click
End Sub

Private Sub Foto1_Click()

    If Val(ZDatos(1, 1)) <> 0 Then

        ZHojaProceso = ZDatos(1, 1)
        ZTerminadoProceso = ZDatos(1, 3)
        ZCantidadProceso = ZDatos(1, 4)
        ZEtapaProceso = ZDatos(1, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
        
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
    
    End If

End Sub

Private Sub Foto2_Click()

    If Val(ZDatos(2, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(2, 1)
        ZTerminadoProceso = ZDatos(2, 3)
        ZCantidadProceso = ZDatos(2, 4)
        ZEtapaProceso = ZDatos(2, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto3_Click()

    If Val(ZDatos(3, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(3, 1)
        ZTerminadoProceso = ZDatos(3, 3)
        ZCantidadProceso = ZDatos(3, 4)
        ZEtapaProceso = ZDatos(3, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto4_Click()

    If Val(ZDatos(4, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(4, 1)
        ZTerminadoProceso = ZDatos(4, 3)
        ZCantidadProceso = ZDatos(4, 4)
        ZEtapaProceso = ZDatos(4, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto5_Click()

    If Val(ZDatos(5, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(5, 1)
        ZTerminadoProceso = ZDatos(5, 3)
        ZCantidadProceso = ZDatos(5, 4)
        ZEtapaProceso = ZDatos(5, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto6_Click()

    If Val(ZDatos(6, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(6, 1)
        ZTerminadoProceso = ZDatos(6, 3)
        ZCantidadProceso = ZDatos(6, 4)
        ZEtapaProceso = ZDatos(6, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto7_Click()

    If Val(ZDatos(7, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(7, 1)
        ZTerminadoProceso = ZDatos(7, 3)
        ZCantidadProceso = ZDatos(7, 4)
        ZEtapaProceso = ZDatos(7, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto8_Click()

    If Val(ZDatos(8, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(8, 1)
        ZTerminadoProceso = ZDatos(8, 3)
        ZCantidadProceso = ZDatos(8, 4)
        ZEtapaProceso = ZDatos(8, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub

Private Sub Foto9_Click()

    If Val(ZDatos(9, 1)) <> 0 Then
    
        ZHojaProceso = ZDatos(9, 1)
        ZTerminadoProceso = ZDatos(9, 3)
        ZCantidadProceso = ZDatos(9, 4)
        ZEtapaProceso = ZDatos(9, 8)
    
        If Val(ZEtapaProceso) = 0 Then
    
            ZEtapaProceso = "1"
            ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
            ZHora = Left$(Time$, 5)
            ZTimer = Int(Timer)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Etapa = " + "'" + ZEtapaProceso + "',"
            ZSql = ZSql + " EstadoHoja = " + "'" + "1" + "',"
            ZSql = ZSql + " FechaInicioEtapa = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicioEtapa = " + "'" + ZHora + "',"
            ZSql = ZSql + " TimerInicioEtapa = " + "'" + Str$(ZTimer) + "',"
            ZSql = ZSql + " FechaInicio = " + "'" + ZFecha + "',"
            ZSql = ZSql + " HoraInicio = " + "'" + ZHora + "',"
            ZSql = ZSql + " Alarma = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaI = " + "'" + "" + "',"
            ZSql = ZSql + " AlarmaII = " + "'" + "" + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZHojaProceso + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        PrgConsultaHojaEnvasado.Hide
        Unload Me
        ZZOrigenProceso = 2
        PrgProcesoNueva.Show
        
    End If

End Sub


Private Sub Proceso_Click()

    Sql1 = "Select *"
    Sql2 = " FROM Operarios"
    Sql3 = " Order by Operarios.Codigo"
    spOperarios = Sql1 + Sql2 + Sql3
    Set rstOperarios = db.OpenRecordset(spOperarios, dbOpenSnapshot, dbSQLPassThrough)
    If rstOperarios.RecordCount > 0 Then
        With rstOperarios
            .MoveFirst
            Do
                If .EOF = False Then
                    ZOperador(rstOperarios!Codigo, 1) = rstOperarios!Codigo
                    ZOperador(rstOperarios!Codigo, 2) = rstOperarios!Descripcion
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstOperarios.Close
    End If
    
    OPEN_FILE_Temperatura0
    OPEN_FILE_Temperatura1
    OPEN_FILE_Temperatura2
    OPEN_FILE_Temperatura3
    OPEN_FILE_Temperatura4
    OPEN_FILE_Temperatura5
    OPEN_FILE_Temperatura6
    OPEN_FILE_Temperatura7
    
    OPEN_FILE_EstadoReactor0
    OPEN_FILE_EstadoReactor1
    OPEN_FILE_EstadoReactor2
    OPEN_FILE_EstadoReactor3
    OPEN_FILE_EstadoReactor4
    OPEN_FILE_EstadoReactor5
    OPEN_FILE_EstadoReactor6
    OPEN_FILE_EstadoReactor7
    

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado0 = 0

    With rstEstadoReactor0
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado0 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado1 = 0

    With rstEstadoReactor1
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado1 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado2 = 0

    With rstEstadoReactor2
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado2 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado3 = 0

    With rstEstadoReactor3
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado3 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado4 = 0

    With rstEstadoReactor4
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado4 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado5 = 0

    With rstEstadoReactor5
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado5 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado6 = 0

    With rstEstadoReactor6
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado6 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0
    WEstado7 = 0

    With rstEstadoReactor7
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                WEstado7 = !Estado
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    
    
    
    
    
    
    
    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura0
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura1.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura1
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura2.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura2
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura3.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura3
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura4.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura4
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura5.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura5
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura6.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0
    WHora = 0

    With rstTemperatura6
        .Index = "Fecha"
        .MoveLast
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                If !Hora > WHora Then
                    ImpreTemperatura7.Caption = !Valor
                    WHora = !Hora
                End If
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With

    WFecha = "31/12/2100"
    WPasa = 0

    With rstTemperatura7
        .Index = "Fecha"
        .Seek "<=", WFecha
        Do
            If .EOF = False Then
                If WPasa = 0 Then
                    WPasa = 1
                    WFecha = !Fecha
                End If
                If !Fecha <> WFecha Then
                    Exit Do
                End If
                ImpreTemperatura8.Caption = !Valor
                .MoveNext
                    Else
                Exit Do
            End If
        Loop
    End With
    
    
    
    Erase ZDatos
    
    
    
        
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.EstadoHoja = 1 and Hoja.Renglon = 1 AND Hoja.TipoEtapa = 0"
    ZSql = ZSql + " Order by Hoja.Hoja"
        
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Select Case Trim(UCase(rstHoja!Equipo))
                        Case "1", "I"
                            Renglon = 1
                        Case "2", "II"
                            Renglon = 2
                        Case "3", "III"
                            Renglon = 3
                        Case "4", "IV"
                            Renglon = 4
                        Case "5", "V"
                            Renglon = 5
                        Case "6", "VI"
                            Renglon = 6
                        Case "7", "VII"
                            Renglon = 7
                        Case "8", "VIII"
                            Renglon = 8
                        Case "9", "IX"
                            Renglon = 9
                        Case Else
                            Renglon = 0
                    End Select

                    ZDatos(Renglon, 1) = Pusing("######", Str$(rstHoja!Hoja))
                    ZDatos(Renglon, 2) = rstHoja!Fecha
                    ZDatos(Renglon, 3) = rstHoja!Producto
                    ZDatos(Renglon, 4) = rstHoja!Teorico
                    ZDatos(Renglon, 5) = ""
                    ZDatos(Renglon, 6) = rstHoja!Equipo
                    ZHoraInicio = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
                    ZDatos(Renglon, 7) = ZHoraInicio
                    ZDatos(Renglon, 8) = IIf(IsNull(rstHoja!etapa), "", rstHoja!etapa)
                    ZHoraInicioEtapa = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
                    ZDatos(Renglon, 9) = ZHoraInicioEtapa
                    ZDatos(Renglon, 10) = ZOperador(rstHoja!Operario, 2)
                    ZZTiempoII = IIf(IsNull(rstHoja!TiempoII), "", rstHoja!TiempoII)
                    ZDatos(Renglon, 11) = ZZTiempoII
                    ZZAlarma = IIf(IsNull(rstHoja!alarma), "", rstHoja!alarma)
                    ZDatos(Renglon, 12) = ZZAlarma
                    ZZAlarmaI = IIf(IsNull(rstHoja!AlarmaI), "", rstHoja!AlarmaI)
                    ZDatos(Renglon, 13) = ZZAlarmaI
                    ZZAlarmaII = IIf(IsNull(rstHoja!AlarmaII), "", rstHoja!AlarmaII)
                    ZDatos(Renglon, 14) = ZZAlarmaII
                    ZTimer = IIf(IsNull(rstHoja!TimerInicioEtapa), "0", rstHoja!TimerInicioEtapa)
                    ZDatos(Renglon, 15) = ZTimer
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    

    
    If WEstado0 = 1 Then
        Foto1.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura1.BackColor = &HC0FFFF
        ImpreTemperatura1.ForeColor = &H404040
            Else
        Foto1.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura1.BackColor = &H404040
        ImpreTemperatura1.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(1, 1)) <> 0 Then
    
        ImpreHoja1.Visible = True
        ImpreProducto1.Visible = True
        ImpreCantidad1.Visible = True
        ImpreOperador1.Visible = True
        ImpreEtapa1.Visible = True
        ImpreHora1.Visible = True
        ImpreTemperatura1.Visible = True
        
        ImpreHoja1.Caption = "Hoja : " + ZDatos(1, 1)
        ImpreProducto1.Caption = ZDatos(1, 3)
        ImpreCantidad1.Caption = "Cantidad : " + ZDatos(1, 4)
        ImpreOperador1.Caption = ZDatos(1, 10)
        ImpreEtapa1.Caption = "Etapa : " + ZDatos(1, 8)
        ImpreHora1.Caption = "Inicio : " + ZDatos(1, 9)
        
        ZTiempoII = Val(ZDatos(1, 11))
        ZAlarma = Trim(ZDatos(1, 12))
        ZAlarmaI = Trim(ZDatos(1, 13))
        ZAlarmaII = Trim(ZDatos(1, 14))
        ZTimer = Val(ZDatos(1, 15))
        ZHoja = ZDatos(1, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(1, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(1, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If

        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura1.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura1.Caption) >= ZZDesdeI And Val(ImpreTemperatura1.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura1.Caption) < ZZDesdeI Or Val(ImpreTemperatura1.Caption) > ZZHastaI Then
                ImpreTemperatura1.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura1.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja1.Visible = False
        ImpreProducto1.Visible = False
        ImpreCantidad1.Visible = False
        ImpreOperador1.Visible = False
        ImpreEtapa1.Visible = False
        ImpreHora1.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    If WEstado1 = 1 Then
        Foto2.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura2.BackColor = &HC0FFFF
        ImpreTemperatura2.ForeColor = &H404040
            Else
        Foto2.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura2.BackColor = &H404040
        ImpreTemperatura2.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(2, 1)) <> 0 Then
    
        ImpreHoja2.Visible = True
        ImpreProducto2.Visible = True
        ImpreCantidad2.Visible = True
        ImpreOperador2.Visible = True
        ImpreEtapa2.Visible = True
        ImpreHora2.Visible = True
        ImpreTemperatura2.Visible = True
        
        ImpreHoja2.Caption = "Hoja : " + ZDatos(2, 1)
        ImpreProducto2.Caption = ZDatos(2, 3)
        ImpreCantidad2.Caption = "Cantidad : " + ZDatos(2, 4)
        ImpreOperador2.Caption = ZDatos(2, 10)
        ImpreEtapa2.Caption = "Etapa : " + ZDatos(2, 8)
        ImpreHora2.Caption = "Inicio : " + ZDatos(2, 9)
        
        ZTiempoII = Val(ZDatos(2, 11))
        ZAlarma = Trim(ZDatos(2, 12))
        ZAlarmaI = Trim(ZDatos(2, 13))
        ZAlarmaII = Trim(ZDatos(2, 14))
        ZTimer = Val(ZDatos(2, 15))
        ZHoja = ZDatos(2, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(2, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(2, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura2.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura2.Caption) >= ZZDesdeI And Val(ImpreTemperatura2.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura2.Caption) < ZZDesdeI Or Val(ImpreTemperatura2.Caption) > ZZHastaI Then
                ImpreTemperatura2.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura2.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja2.Visible = False
        ImpreProducto2.Visible = False
        ImpreCantidad2.Visible = False
        ImpreOperador2.Visible = False
        ImpreEtapa2.Visible = False
        ImpreHora2.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    If WEstado2 = 1 Then
        Foto3.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura3.BackColor = &HC0FFFF
        ImpreTemperatura3.ForeColor = &H404040
            Else
        Foto3.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura3.BackColor = &H404040
        ImpreTemperatura3.ForeColor = &HFFFFFF
    End If
    
    If Val(ZDatos(3, 1)) <> 0 Then
        
        ImpreHoja3.Visible = True
        ImpreProducto3.Visible = True
        ImpreCantidad3.Visible = True
        ImpreOperador3.Visible = True
        ImpreEtapa3.Visible = True
        ImpreHora3.Visible = True
        ImpreTemperatura3.Visible = True
        
        ImpreHoja3.Caption = "Hoja : " + ZDatos(3, 1)
        ImpreProducto3.Caption = ZDatos(3, 3)
        ImpreCantidad3.Caption = "Cantidad : " + ZDatos(3, 4)
        ImpreOperador3.Caption = ZDatos(3, 10)
        ImpreEtapa3.Caption = "Etapa : " + ZDatos(3, 8)
        ImpreHora3.Caption = "Inicio : " + ZDatos(3, 9)
        
        ZTiempoII = Val(ZDatos(3, 11))
        ZAlarma = Trim(ZDatos(3, 12))
        ZAlarmaI = Trim(ZDatos(3, 13))
        ZAlarmaII = Trim(ZDatos(3, 14))
        ZTimer = Val(ZDatos(3, 15))
        ZHoja = ZDatos(3, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(3, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(3, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura3.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura3.Caption) >= ZZDesdeI And Val(ImpreTemperatura3.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura3.Caption) < ZZDesdeI Or Val(ImpreTemperatura3.Caption) > ZZHastaI Then
                ImpreTemperatura3.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura3.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja3.Visible = False
        ImpreProducto3.Visible = False
        ImpreCantidad3.Visible = False
        ImpreOperador3.Visible = False
        ImpreEtapa3.Visible = False
        ImpreHora3.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    If WEstado3 = 1 Then
        Foto4.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura4.BackColor = &HC0FFFF
        ImpreTemperatura4.ForeColor = &H404040
            Else
        Foto4.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura4.BackColor = &H404040
        ImpreTemperatura4.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(4, 1)) <> 0 Then
    
        ImpreHoja4.Visible = True
        ImpreProducto4.Visible = True
        ImpreCantidad4.Visible = True
        ImpreOperador4.Visible = True
        ImpreEtapa4.Visible = True
        ImpreHora4.Visible = True
        ImpreTemperatura4.Visible = True
        
        ImpreHoja4.Caption = "Hoja : " + ZDatos(4, 1)
        ImpreProducto4.Caption = ZDatos(4, 3)
        ImpreCantidad4.Caption = "Cantidad : " + ZDatos(4, 4)
        ImpreOperador4.Caption = ZDatos(4, 10)
        ImpreEtapa4.Caption = "Etapa : " + ZDatos(4, 8)
        ImpreHora4.Caption = "Inicio : " + ZDatos(4, 9)
        
        ZTiempoII = Val(ZDatos(4, 11))
        ZAlarma = Trim(ZDatos(4, 12))
        ZAlarmaI = Trim(ZDatos(4, 13))
        ZAlarmaII = Trim(ZDatos(4, 14))
        ZTimer = Val(ZDatos(4, 15))
        ZHoja = ZDatos(4, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(4, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(4, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura4.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura4.Caption) >= ZZDesdeI And Val(ImpreTemperatura4.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura4.Caption) < ZZDesdeI Or Val(ImpreTemperatura4.Caption) > ZZHastaI Then
                ImpreTemperatura4.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura4.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja4.Visible = False
        ImpreProducto4.Visible = False
        ImpreCantidad4.Visible = False
        ImpreOperador4.Visible = False
        ImpreEtapa4.Visible = False
        ImpreHora4.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    If WEstado4 = 1 Then
        Foto5.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura5.BackColor = &HC0FFFF
        ImpreTemperatura5.ForeColor = &H404040
            Else
        Foto5.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura5.BackColor = &H404040
        ImpreTemperatura5.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(5, 1)) <> 0 Then
    
        ImpreHoja5.Visible = True
        ImpreProducto5.Visible = True
        ImpreCantidad5.Visible = True
        ImpreOperador5.Visible = True
        ImpreEtapa5.Visible = True
        ImpreHora5.Visible = True
        ImpreTemperatura5.Visible = True
        
        ImpreHoja5.Caption = "Hoja : " + ZDatos(5, 1)
        ImpreProducto5.Caption = ZDatos(5, 3)
        ImpreCantidad5.Caption = "Cantidad : " + ZDatos(5, 4)
        ImpreOperador5.Caption = ZDatos(5, 10)
        ImpreEtapa5.Caption = "Etapa : " + ZDatos(5, 8)
        ImpreHora5.Caption = "Inicio : " + ZDatos(5, 9)
        
        ZTiempoII = Val(ZDatos(5, 11))
        ZAlarma = Trim(ZDatos(5, 12))
        ZAlarmaI = Trim(ZDatos(5, 13))
        ZAlarmaII = Trim(ZDatos(5, 14))
        ZTimer = Val(ZDatos(5, 15))
        ZHoja = ZDatos(5, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(5, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(5, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura5.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura5.Caption) >= ZZDesdeI And Val(ImpreTemperatura5.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura5.Caption) < ZZDesdeI Or Val(ImpreTemperatura5.Caption) > ZZHastaI Then
                ImpreTemperatura5.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura5.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja5.Visible = False
        ImpreProducto5.Visible = False
        ImpreCantidad5.Visible = False
        ImpreOperador5.Visible = False
        ImpreEtapa5.Visible = False
        ImpreHora5.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    If WEstado5 = 1 Then
        Foto6.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura6.BackColor = &HC0FFFF
        ImpreTemperatura6.ForeColor = &H404040
            Else
        Foto6.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura6.BackColor = &H404040
        ImpreTemperatura6.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(6, 1)) <> 0 Then
    
        ImpreHoja6.Visible = True
        ImpreProducto6.Visible = True
        ImpreCantidad6.Visible = True
        ImpreOperador6.Visible = True
        ImpreEtapa6.Visible = True
        ImpreHora6.Visible = True
        ImpreTemperatura6.Visible = True
        
        ImpreHoja6.Caption = "Hoja : " + ZDatos(6, 1)
        ImpreProducto6.Caption = ZDatos(6, 3)
        ImpreCantidad6.Caption = "Cantidad : " + ZDatos(6, 4)
        ImpreOperador6.Caption = ZDatos(6, 10)
        ImpreEtapa6.Caption = "Etapa : " + ZDatos(6, 8)
        ImpreHora6.Caption = "Inicio : " + ZDatos(6, 9)
        
        ZTiempoII = Val(ZDatos(6, 11))
        ZAlarma = Trim(ZDatos(6, 12))
        ZAlarmaI = Trim(ZDatos(6, 13))
        ZAlarmaII = Trim(ZDatos(6, 14))
        ZTimer = Val(ZDatos(6, 15))
        ZHoja = ZDatos(6, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(6, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(6, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura6.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura6.Caption) >= ZZDesdeI And Val(ImpreTemperatura6.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura6.Caption) < ZZDesdeI Or Val(ImpreTemperatura6.Caption) > ZZHastaI Then
                ImpreTemperatura6.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura6.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja6.Visible = False
        ImpreProducto6.Visible = False
        ImpreCantidad6.Visible = False
        ImpreOperador6.Visible = False
        ImpreEtapa6.Visible = False
        ImpreHora6.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    
    If WEstado6 = 1 Then
        Foto7.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura7.BackColor = &HC0FFFF
        ImpreTemperatura7.ForeColor = &H404040
            Else
        Foto7.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura7.BackColor = &H404040
        ImpreTemperatura7.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(7, 1)) <> 0 Then
    
        ImpreHoja7.Visible = True
        ImpreProducto7.Visible = True
        ImpreCantidad7.Visible = True
        ImpreOperador7.Visible = True
        ImpreEtapa7.Visible = True
        ImpreHora7.Visible = True
        ImpreTemperatura7.Visible = True
        
        ImpreHoja7.Caption = "Hoja : " + ZDatos(7, 1)
        ImpreProducto7.Caption = ZDatos(7, 3)
        ImpreCantidad7.Caption = "Cantidad : " + ZDatos(7, 4)
        ImpreOperador7.Caption = ZDatos(7, 10)
        ImpreEtapa7.Caption = "Etapa : " + ZDatos(7, 8)
        ImpreHora7.Caption = "Inicio : " + ZDatos(7, 9)
        
        ZTiempoII = Val(ZDatos(7, 11))
        ZAlarma = Trim(ZDatos(7, 12))
        ZAlarmaI = Trim(ZDatos(7, 13))
        ZAlarmaII = Trim(ZDatos(7, 14))
        ZTimer = Val(ZDatos(7, 15))
        ZHoja = ZDatos(7, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(7, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(7, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura7.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura7.Caption) >= ZZDesdeI And Val(ImpreTemperatura7.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura7.Caption) < ZZDesdeI Or Val(ImpreTemperatura7.Caption) > ZZHastaI Then
                ImpreTemperatura7.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura7.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja7.Visible = False
        ImpreProducto7.Visible = False
        ImpreCantidad7.Visible = False
        ImpreOperador7.Visible = False
        ImpreEtapa7.Visible = False
        ImpreHora7.Visible = False
        
    End If
    
    
    
    
    
    
    
    
    
    
    If WEstado7 = 1 Then
        Foto8.Picture = LoadPicture("Tanqueverde.jpg")
        ImpreTemperatura8.BackColor = &HC0FFFF
        ImpreTemperatura8.ForeColor = &H404040
            Else
        Foto8.Picture = LoadPicture("Tanquenegro.jpg")
        ImpreTemperatura8.BackColor = &H404040
        ImpreTemperatura8.ForeColor = &HFFFFFF
    End If
        
    If Val(ZDatos(8, 1)) <> 0 Then
    
        ImpreHoja8.Visible = True
        ImpreProducto8.Visible = True
        ImpreCantidad8.Visible = True
        ImpreOperador8.Visible = True
        ImpreEtapa8.Visible = True
        ImpreHora8.Visible = True
        ImpreTemperatura8.Visible = True
        
        ImpreHoja8.Caption = "Hoja : " + ZDatos(8, 1)
        ImpreProducto8.Caption = ZDatos(8, 3)
        ImpreCantidad8.Caption = "Cantidad : " + ZDatos(8, 4)
        ImpreOperador8.Caption = ZDatos(8, 10)
        ImpreEtapa8.Caption = "Etapa : " + ZDatos(8, 8)
        ImpreHora8.Caption = "Inicio : " + ZDatos(8, 9)
        
        ZTiempoII = Val(ZDatos(8, 11))
        ZAlarma = Trim(ZDatos(8, 12))
        ZAlarmaI = Trim(ZDatos(8, 13))
        ZAlarmaII = Trim(ZDatos(8, 14))
        ZTimer = Val(ZDatos(8, 15))
        ZHoja = ZDatos(8, 1)
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM ProduIII"
        ZSql = ZSql + " Where ProduIII.Terminado = " + "'" + ZDatos(8, 3) + "'"
        ZSql = ZSql + " and ProduIII.Etapa = " + "'" + ZDatos(8, 8) + "'"
        rsProduIII = ZSql
        Set rstProduIII = db.OpenRecordset(rsProduIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstProduIII.RecordCount > 0 Then
        
            ZZControlI = rstProduIII!ControlI
            ZZDesdeI = rstProduIII!TemperaturaI
            ZZHastaI = rstProduIII!TemperaturaII
            ZZTiempoI = rstProduIII!Tiempo
            
            ZZControlII = rstProduIII!ControlII
            ZZTiempoII = rstProduIII!TiempoII
            
            rstProduIII.Close
        End If
        
        ZTimeractual = Int(Timer)
        ZSegundos = ZTimeractual - ZTimer
        ZTiempoI = Int(ZSegundos / 60)
    
        If ZZControlII = 1 And ZTiempoI > ZZTiempoII Then
            If Val(ImpreTemperatura8.Caption) < ZZDesdeI Then
                If ZAlarma <> "S" Then
                    Rem m$ = "Ha pasado el tiempo estimado para alacanzar la temperatura establecida en esta etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarma.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " Alarma = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
        If ZTiempoII = 0 Then
            If Val(ImpreTemperatura8.Caption) >= ZZDesdeI And Val(ImpreTemperatura8.Caption) < ZZHastaI Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Hoja SET "
                ZSql = ZSql + " TiempoII = " + "'" + Str$(ZTiempoI) + "'"
                ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                spHoja = ZSql
                Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    
        ZTiempoIII = ZTiempoI - ZTiempoII
    
        If ZZControlI = 1 Then
            If ZTiempoIII > ZZTiempoI Then
                If ZAlarmaII <> "S" Then
                    Rem m$ = "Se ha excedido el tiempo establecido para la etapa"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    Rem WAlarmaII.Value = 1
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaII = " + "'" + "S" + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
    
        If ZZControlI = 1 And ZTiempoII <> 0 Then
            If Val(ImpreTemperatura8.Caption) < ZZDesdeI Or Val(ImpreTemperatura8.Caption) > ZZHastaI Then
                ImpreTemperatura9.BackColor = &HFF&
                If ZAlarmaI <> "S" Then
                    Rem m$ = "La temperatura no se encuentra dentro del rango establecido"
                    Rem G% = MsgBox(m$, 16, "Carga de Procesos")
                    WAlarmaI.Value = 1
                    WAlarmaITiempo.Text = ZTiempoIII.Text
                    WAlarmaITempe.Text = Temperatura.Text
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Hoja SET "
                    ZSql = ZSql + " AlarmaI = " + "'" + "S" + "',"
                    ZSql = ZSql + " AlarmaITiempo = " + "'" + Str$(ZTiempoIII) + "',"
                    ZSql = ZSql + " AlarmaITempe = " + "'" + ImpreTemperatura8.Caption + "'"
                    ZSql = ZSql + " Where Hoja = " + "'" + ZHoja + "'"
                    spHoja = ZSql
                    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
                End If
            End If
        End If
        
            Else
        
        ImpreHoja8.Visible = False
        ImpreProducto8.Visible = False
        ImpreCantidad8.Visible = False
        ImpreOperador8.Visible = False
        ImpreEtapa8.Visible = False
        ImpreHora8.Visible = False
        
    End If
    
End Sub

Private Sub ImpreOperador1_DblClick()

    ZZCambiaOperario = 1
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador2_DblClick()

    ZZCambiaOperario = 2
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador3_DblClick()

    ZZCambiaOperario = 3
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador4_DblClick()

    ZZCambiaOperario = 4
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador5_DblClick()

    ZZCambiaOperario = 5
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador6_DblClick()

    ZZCambiaOperario = 6
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador7_DblClick()

    ZZCambiaOperario = 7
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador8_DblClick()

    ZZCambiaOperario = 8
    
    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ImpreOperador9_DblClick()

    ZZCambiaOperario = 9

    ListaCambiaOperador.Clear
    
    For Ciclo = 1 To 100
        ListaCambiaOperador.AddItem ZOperador(Ciclo, 2)
    Next Ciclo
    
    CambiaOperador.Height = 4695
    CambiaOperador.Left = 3480
    CambiaOperador.Top = 1080
    CambiaOperador.Width = 4215
    
    CambiaOperador.Visible = True

End Sub

Private Sub ListaCambiaOperador_DblClick()

    ZZLugarOperario = ListaCambiaOperador.ListIndex + 1

    Select Case ZZCambiaOperario
        Case 1
            ImpreOperador1.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja1.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 2
            ImpreOperador2.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja2.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 3
            ImpreOperador3.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja3.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 4
            ImpreOperador4.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja4.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 5
            ImpreOperador5.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja5.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 6
            ImpreOperador6.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja6.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 7
            ImpreOperador7.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja7.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 8
            ImpreOperador8.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja8.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case 9
            ImpreOperador9.Caption = ListaCambiaOperador.Text
            
            ZZHoja = Mid$(ImpreHoja9.Caption, 8, 6)
            
            ZSql = ""
            ZSql = ZSql + "UPDATE Hoja SET "
            ZSql = ZSql + " Operario = " + "'" + ZOperador(ZZLugarOperario, 1) + "'"
            ZSql = ZSql + " Where Hoja = " + "'" + ZZHoja + "'"
            spHoja = ZSql
            Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
            
        Case Else
        
    End Select
    
    CambiaOperador.Visible = False

End Sub

Private Sub Form_Load()
    Timer1.Interval = 30000
End Sub

Private Sub Timer1_Timer()
    Call Proceso_Click
End Sub
