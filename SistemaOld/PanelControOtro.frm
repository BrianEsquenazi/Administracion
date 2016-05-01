VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgPanelControl 
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
      Left            =   9480
      Top             =   7680
   End
   Begin VB.Frame Box9 
      Height          =   2415
      Left            =   8040
      TabIndex        =   8
      Top             =   5160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.Image Stop9 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 9"
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label ImpreTemperatura9 
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
         TabIndex        =   69
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label ImpreHoja9 
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
         TabIndex        =   38
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label ImpreProducto9 
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
         TabIndex        =   37
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label ImpreCantidad9 
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
         TabIndex        =   36
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label ImpreOperador9 
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
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreEtapa9 
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
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label ImpreHora9 
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
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto9 
         Height          =   855
         Left            =   240
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box7 
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   5160
      Width           =   3615
      Begin VB.Image Stop7 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 7"
         Height          =   255
         Left            =   240
         TabIndex        =   77
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
         TabIndex        =   66
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
   Begin VB.Frame Box8 
      Height          =   2415
      Left            =   4200
      TabIndex        =   6
      Top             =   5160
      Width           =   3615
      Begin VB.Image Stop8 
         Height          =   615
         Left            =   480
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 8"
         Height          =   255
         Left            =   360
         TabIndex        =   76
         Top             =   240
         Width           =   1095
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
         TabIndex        =   44
         Top             =   240
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
         TabIndex        =   43
         Top             =   600
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
         TabIndex        =   42
         Top             =   960
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
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   40
         Top             =   1680
         Visible         =   0   'False
         Width           =   1935
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
         TabIndex        =   39
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Image Foto8 
         Height          =   855
         Left            =   360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Box4 
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   2640
      Width           =   3615
      Begin VB.Image Stop4 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 4"
         Height          =   255
         Left            =   240
         TabIndex        =   74
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
         TabIndex        =   64
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
      Left            =   4200
      TabIndex        =   4
      Top             =   2640
      Width           =   3615
      Begin VB.Image Stop5 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 5"
         Height          =   255
         Left            =   240
         TabIndex        =   75
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
         TabIndex        =   68
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
      Left            =   8040
      TabIndex        =   3
      Top             =   2640
      Width           =   3615
      Begin VB.Image Stop6 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 6"
         Height          =   255
         Left            =   240
         TabIndex        =   79
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
         TabIndex        =   70
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
      Left            =   8040
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      Begin VB.Image Stop3 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 3"
         Height          =   255
         Left            =   240
         TabIndex        =   78
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
         TabIndex        =   65
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
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Image Stop2 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 2"
         Height          =   255
         Left            =   240
         TabIndex        =   73
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
         TabIndex        =   63
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Image Stop1 
         Height          =   615
         Left            =   360
         Stretch         =   -1  'True
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reactor 1"
         Height          =   255
         Left            =   240
         TabIndex        =   72
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
         TabIndex        =   71
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label ImpreTemperatura1 
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
      MouseIcon       =   "PanelControOtro.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "PanelControOtro.frx":030A
      ToolTipText     =   "Salida"
      Top             =   7680
      Width           =   480
   End
End
Attribute VB_Name = "PrgPanelControl"
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
    Close
    End
End Sub

Private Sub Estado_click()
    Rem Call Proceso_Click
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Operario_Keypress(KeyAscii As Integer)
    Call Proceso_Click
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
    
        
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.EstadoHoja = 1 and Hoja.Renglon = 1"
    ZSql = ZSql + " Order by Hoja.Hoja"
        
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Select Case Trim(UCase(rstHoja!EquipoII))
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
                    Rem ZDatos(Renglon, 6) = rstHoja!Equipo
                    Rem ZHoraInicio = IIf(IsNull(rstHoja!HoraInicio), "", rstHoja!HoraInicio)
                    Rem ZDatos(Renglon, 7) = ZHoraInicio
                    ZDatos(Renglon, 8) = IIf(IsNull(rstHoja!Etapa), "", rstHoja!Etapa)
                    ZHoraInicioEtapa = IIf(IsNull(rstHoja!HoraInicioEtapa), "", rstHoja!HoraInicioEtapa)
                    ZDatos(Renglon, 9) = ZHoraInicioEtapa
                    ZDatos(Renglon, 10) = ZOperador(rstHoja!Operario, 2)
                    Rem  ZZTiempoII = IIf(IsNull(rstHoja!TiempoII), "", rstHoja!TiempoII)
                    Rem ZDatos(Renglon, 11) = ZZTiempoII
                    Rem ZZAlarma = IIf(IsNull(rstHoja!alarma), "", rstHoja!alarma)
                    Rem ZDatos(Renglon, 12) = ZZAlarma
                    Rem ZZAlarmaI = IIf(IsNull(rstHoja!AlarmaI), "", rstHoja!AlarmaI)
                    Rem ZDatos(Renglon, 13) = ZZAlarmaI
                    Rem ZZAlarmaII = IIf(IsNull(rstHoja!AlarmaII), "", rstHoja!AlarmaII)
                    Rem ZDatos(Renglon, 14) = ZZAlarmaII
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    If Val(ZDatos(1, 1)) <> 0 Then
    
        
        ImpreHoja1.Visible = True
        ImpreProducto1.Visible = True
        ImpreCantidad1.Visible = True
        ImpreOperador1.Visible = True
        Rem ImpreEtapa1.Visible = True
        Rem ImpreHora1.Visible = True
        ImpreTemperatura1.Visible = True
        
        ImpreHoja1.Caption = "Hoja : " + ZDatos(1, 1)
        ImpreProducto1.Caption = ZDatos(1, 3)
        ImpreCantidad1.Caption = "Cantidad : " + ZDatos(1, 4)
        ImpreOperador1.Caption = ZDatos(1, 10)
        ImpreEtapa1.Caption = "Etapa : " + ZDatos(1, 8)
        ImpreHora1.Caption = "Inicio : " + ZDatos(1, 9)
        
            Else
        
        ImpreHoja1.Visible = False
        ImpreProducto1.Visible = False
        ImpreCantidad1.Visible = False
        ImpreOperador1.Visible = False
        ImpreEtapa1.Visible = False
        ImpreHora1.Visible = False
        
    End If
    
    If Val(ZDatos(2, 1)) <> 0 Then
    
        ImpreHoja2.Visible = True
        ImpreProducto2.Visible = True
        ImpreCantidad2.Visible = True
        ImpreOperador2.Visible = True
        Rem ImpreEtapa2.Visible = True
        Rem ImpreHora2.Visible = True
        ImpreTemperatura2.Visible = True
        
        ImpreHoja2.Caption = "Hoja : " + ZDatos(2, 1)
        ImpreProducto2.Caption = ZDatos(2, 3)
        ImpreCantidad2.Caption = "Cantidad : " + ZDatos(2, 4)
        ImpreOperador2.Caption = ZDatos(2, 10)
        ImpreEtapa2.Caption = "Etapa : " + ZDatos(2, 8)
        ImpreHora2.Caption = "Inicio : " + ZDatos(2, 9)
        
            Else
            
        ImpreHoja2.Visible = False
        ImpreProducto2.Visible = False
        ImpreCantidad2.Visible = False
        ImpreOperador2.Visible = False
        ImpreEtapa2.Visible = False
        ImpreHora2.Visible = False
        
    End If
    
    If Val(ZDatos(3, 1)) <> 0 Then
    
        ImpreHoja3.Visible = True
        ImpreProducto3.Visible = True
        ImpreCantidad3.Visible = True
        ImpreOperador3.Visible = True
        Rem ImpreEtapa3.Visible = True
        Rem ImpreHora3.Visible = True
        ImpreTemperatura3.Visible = True
        
        ImpreHoja3.Caption = "Hoja : " + ZDatos(3, 1)
        ImpreProducto3.Caption = ZDatos(3, 3)
        ImpreCantidad3.Caption = "Cantidad : " + ZDatos(3, 4)
        ImpreOperador3.Caption = ZDatos(3, 10)
        ImpreEtapa3.Caption = "Etapa : " + ZDatos(3, 8)
        ImpreHora3.Caption = "Inicio : " + ZDatos(3, 9)
        
            Else
            
        ImpreHoja3.Visible = False
        ImpreProducto3.Visible = False
        ImpreCantidad3.Visible = False
        ImpreOperador3.Visible = False
        ImpreEtapa3.Visible = False
        ImpreHora3.Visible = False
        
    End If
    
    If Val(ZDatos(4, 1)) <> 0 Then
    
        ImpreHoja4.Visible = True
        ImpreProducto4.Visible = True
        ImpreCantidad4.Visible = True
        ImpreOperador4.Visible = True
        Rem ImpreEtapa4.Visible = True
        Rem ImpreHora4.Visible = True
        ImpreTemperatura4.Visible = True
        
        ImpreHoja4.Caption = "Hoja : " + ZDatos(4, 1)
        ImpreProducto4.Caption = ZDatos(4, 3)
        ImpreCantidad4.Caption = "Cantidad : " + ZDatos(4, 4)
        ImpreOperador4.Caption = ZDatos(4, 10)
        ImpreEtapa4.Caption = "Etapa : " + ZDatos(4, 8)
        ImpreHora4.Caption = "Inicio : " + ZDatos(4, 9)
        
            Else
            
        ImpreHoja4.Visible = False
        ImpreProducto4.Visible = False
        ImpreCantidad4.Visible = False
        ImpreOperador4.Visible = False
        ImpreEtapa4.Visible = False
        ImpreHora4.Visible = False
        
    End If
    
    If Val(ZDatos(5, 1)) <> 0 Then
    
        ImpreHoja5.Visible = True
        ImpreProducto5.Visible = True
        ImpreCantidad5.Visible = True
        ImpreOperador5.Visible = True
        Rem ImpreEtapa5.Visible = True
        Rem ImpreHora5.Visible = True
        ImpreTemperatura5.Visible = True
        
        ImpreHoja5.Caption = "Hoja : " + ZDatos(5, 1)
        ImpreProducto5.Caption = ZDatos(5, 3)
        ImpreCantidad5.Caption = "Cantidad : " + ZDatos(5, 4)
        ImpreOperador5.Caption = ZDatos(5, 10)
        ImpreEtapa5.Caption = "Etapa : " + ZDatos(5, 8)
        ImpreHora5.Caption = "Inicio : " + ZDatos(5, 9)
        
            Else
            
        ImpreHoja5.Visible = False
        ImpreProducto5.Visible = False
        ImpreCantidad5.Visible = False
        ImpreOperador5.Visible = False
        ImpreEtapa5.Visible = False
        ImpreHora5.Visible = False
        
    End If
    
    If Val(ZDatos(6, 1)) <> 0 Then
    
        ImpreHoja6.Visible = True
        ImpreProducto6.Visible = True
        ImpreCantidad6.Visible = True
        ImpreOperador6.Visible = True
        Rem ImpreEtapa6.Visible = True
        Rem ImpreHora6.Visible = True
        ImpreTemperatura6.Visible = True
        
        ImpreHoja6.Caption = "Hoja : " + ZDatos(6, 1)
        ImpreProducto6.Caption = ZDatos(6, 3)
        ImpreCantidad6.Caption = "Cantidad : " + ZDatos(6, 4)
        ImpreOperador6.Caption = ZDatos(6, 10)
        ImpreEtapa6.Caption = "Etapa : " + ZDatos(6, 8)
        ImpreHora6.Caption = "Inicio : " + ZDatos(6, 9)
        
            Else
            
        ImpreHoja6.Visible = False
        ImpreProducto6.Visible = False
        ImpreCantidad6.Visible = False
        ImpreOperador6.Visible = False
        ImpreEtapa6.Visible = False
        ImpreHora6.Visible = False
        
    End If
    
    If Val(ZDatos(7, 1)) <> 0 Then
    
        ImpreHoja7.Visible = True
        ImpreProducto7.Visible = True
        ImpreCantidad7.Visible = True
        ImpreOperador7.Visible = True
        Rem ImpreEtapa7.Visible = True
        Rem ImpreHora7.Visible = True
        ImpreTemperatura7.Visible = True
        
        ImpreHoja7.Caption = "Hoja : " + ZDatos(7, 1)
        ImpreProducto7.Caption = ZDatos(7, 3)
        ImpreCantidad7.Caption = "Cantidad : " + ZDatos(7, 4)
        ImpreOperador7.Caption = ZDatos(7, 10)
        ImpreEtapa7.Caption = "Etapa : " + ZDatos(7, 8)
        ImpreHora7.Caption = "Inicio : " + ZDatos(7, 9)
        
            Else
            
        ImpreHoja7.Visible = False
        ImpreProducto7.Visible = False
        ImpreCantidad7.Visible = False
        ImpreOperador7.Visible = False
        ImpreEtapa7.Visible = False
        ImpreHora7.Visible = False
        
    End If
    
    If Val(ZDatos(8, 1)) <> 0 Then
    
        ImpreHoja8.Visible = True
        ImpreProducto8.Visible = True
        ImpreCantidad8.Visible = True
        ImpreOperador8.Visible = True
        Rem ImpreEtapa8.Visible = True
        Rem ImpreHora8.Visible = True
        ImpreTemperatura8.Visible = True
        
        ImpreHoja8.Caption = "Hoja : " + ZDatos(8, 1)
        ImpreProducto8.Caption = ZDatos(8, 3)
        ImpreCantidad8.Caption = "Cantidad : " + ZDatos(8, 4)
        ImpreOperador8.Caption = ZDatos(8, 10)
        ImpreEtapa8.Caption = "Etapa : " + ZDatos(8, 8)
        ImpreHora8.Caption = "Inicio : " + ZDatos(8, 9)
        
            Else
            
        ImpreHoja8.Visible = False
        ImpreProducto8.Visible = False
        ImpreCantidad8.Visible = False
        ImpreOperador8.Visible = False
        ImpreEtapa8.Visible = False
        ImpreHora8.Visible = False
        
    End If
    
    If Val(ZDatos(9, 1)) <> 0 Then
    
        ImpreHoja9.Visible = True
        ImpreProducto9.Visible = True
        ImpreCantidad9.Visible = True
        ImpreOperador9.Visible = True
        Rem ImpreEtapa9.Visible = True
        Rem ImpreHora9.Visible = True
        ImpreTemperatura9.Visible = True
        
        ImpreHoja9.Caption = "Hoja : " + ZDatos(9, 1)
        ImpreProducto9.Caption = ZDatos(9, 3)
        ImpreCantidad9.Caption = "Cantidad : " + ZDatos(9, 4)
        ImpreOperador9.Caption = ZDatos(9, 10)
        ImpreEtapa9.Caption = "Etapa : " + ZDatos(9, 8)
        ImpreHora9.Caption = "Inicio : " + ZDatos(9, 9)
        
            Else
            
        ImpreHoja9.Visible = False
        ImpreProducto9.Visible = False
        ImpreCantidad9.Visible = False
        ImpreOperador9.Visible = False
        ImpreEtapa9.Visible = False
        ImpreHora9.Visible = False
        
    End If
    
    Call Timer1_Timer
    
End Sub

Private Sub Form_Load()
    Timer1.Interval = 10000
End Sub

Private Sub Timer1_Timer()

    On Error GoTo WError


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
    

    WEstado0 = 1
    With rstEstadoReactor0
        .Index = "Fecha"
        .MoveLast
        WEstado0 = !Estado
    End With

    WEstado1 = 1
    With rstEstadoReactor1
        .Index = "Fecha"
        .MoveLast
        WEstado1 = !Estado
    End With

    WEstado2 = 1
    With rstEstadoReactor2
        .Index = "Fecha"
        .MoveLast
        WEstado2 = !Estado
    End With

    WEstado3 = 1
    With rstEstadoReactor3
        .Index = "Fecha"
        .MoveLast
        WEstado3 = !Estado
    End With

    WEstado4 = 1
    With rstEstadoReactor4
        .Index = "Fecha"
        .MoveLast
        WEstado4 = !Estado
    End With

    WEstado5 = 1
    With rstEstadoReactor5
        .Index = "Fecha"
        .MoveLast
        WEstado5 = !Estado
    End With

    WEstado6 = 1
    With rstEstadoReactor6
        .Index = "Fecha"
        .MoveLast
        WEstado6 = !Estado
    End With

    WEstado7 = 1
    With rstEstadoReactor7
        .Index = "Fecha"
        .MoveLast
        WEstado7 = !Estado
    End With
    
    If WEstado0 = 0 Then
        Foto1.Picture = LoadPicture("Tanqueverde.jpg")
        Stop1.Picture = LoadPicture("")
        ImpreTemperatura1.BackColor = &HC0FFFF
            Else
        Foto1.Picture = LoadPicture("Tanquenegro.jpg")
        Stop1.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura1.BackColor = &HC0FFFF
        Rem ImpreTemperatura1.BackColor = &HFF&
    End If
    
    If WEstado1 = 0 Then
        Foto2.Picture = LoadPicture("Tanqueverde.jpg")
        Stop2.Picture = LoadPicture("")
        ImpreTemperatura2.BackColor = &HC0FFFF
            Else
        Foto2.Picture = LoadPicture("Tanquenegro.jpg")
        Stop2.Picture = LoadPicture("stop.jpg")
        Rem ImpreTemperatura2.BackColor = &HFF&
        ImpreTemperatura2.BackColor = &HC0FFFF
    End If
    
    If WEstado2 = 0 Then
        Foto3.Picture = LoadPicture("Tanqueverde.jpg")
        Stop3.Picture = LoadPicture("")
        ImpreTemperatura3.BackColor = &HC0FFFF
            Else
        Foto3.Picture = LoadPicture("Tanquenegro.jpg")
        Stop3.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura3.BackColor = &HC0FFFF
        Rem ImpreTemperatura3.BackColor = &HFF&
    End If
    
    If WEstado3 = 0 Then
        Foto4.Picture = LoadPicture("Tanqueverde.jpg")
        Stop4.Picture = LoadPicture("")
        ImpreTemperatura4.BackColor = &HC0FFFF
            Else
        Foto4.Picture = LoadPicture("Tanquenegro.jpg")
        Stop4.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura4.BackColor = &HC0FFFF
        Rem ImpreTemperatura4.BackColor = &HFF&
    End If
    
    If WEstado4 = 0 Then
        Foto5.Picture = LoadPicture("Tanqueverde.jpg")
        Stop5.Picture = LoadPicture("")
        ImpreTemperatura5.BackColor = &HC0FFFF
            Else
        Foto5.Picture = LoadPicture("Tanquenegro.jpg")
        Stop5.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura5.BackColor = &HC0FFFF
        Rem ImpreTemperatura5.BackColor = &HFF&
    End If
    
    If WEstado5 = 0 Then
        Foto6.Picture = LoadPicture("Tanqueverde.jpg")
        Stop6.Picture = LoadPicture("")
        ImpreTemperatura6.BackColor = &HC0FFFF
            Else
        Foto6.Picture = LoadPicture("Tanquenegro.jpg")
        Stop6.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura6.BackColor = &HC0FFFF
        Rem ImpreTemperatura6.BackColor = &HFF&
    End If
    
    If WEstado6 = 0 Then
        Foto7.Picture = LoadPicture("Tanqueverde.jpg")
        Stop7.Picture = LoadPicture("")
        ImpreTemperatura7.BackColor = &HC0FFFF
            Else
        Foto7.Picture = LoadPicture("Tanquenegro.jpg")
        Stop7.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura7.BackColor = &HC0FFFF
        Rem ImpreTemperatura7.BackColor = &HFF&
    End If
    
    If WEstado7 = 0 Then
        Foto8.Picture = LoadPicture("Tanqueverde.jpg")
        Stop8.Picture = LoadPicture("")
        ImpreTemperatura8.BackColor = &HC0FFFF
            Else
        Foto8.Picture = LoadPicture("Tanquenegro.jpg")
        Stop8.Picture = LoadPicture("stop.jpg")
        ImpreTemperatura8.BackColor = &HC0FFFF
        Rem ImpreTemperatura8.BackColor = &HFF&
    End If
    
    
    With rstTemperatura0
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura1.Caption = !Valor
    End With
    
    With rstTemperatura1
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura2.Caption = !Valor
    End With
    
    With rstTemperatura2
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura3.Caption = !Valor
    End With
    
    With rstTemperatura3
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura4.Caption = !Valor
    End With
    
    With rstTemperatura4
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura5.Caption = !Valor
    End With
    
    With rstTemperatura5
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura6.Caption = !Valor
    End With
    
    With rstTemperatura6
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura7.Caption = !Valor
    End With
    
    With rstTemperatura7
        .Index = "Fecha"
        .MoveLast
        ImpreTemperatura8.Caption = !Valor
    End With
    
    Exit Sub
    
WError:
    Resume Next
    
    
End Sub



