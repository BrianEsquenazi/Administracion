VERSION 5.00
Begin VB.Form PrgPrueArtRango 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame NroLote 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.CommandButton FinNroLote 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   5520
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "72.700   a 72.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   52
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "580.000 a 584.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   51
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "540.000 a 559.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   50
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "SURFACTAN  VII"
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
         TabIndex        =   49
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "72.400   a 72.699"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   48
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "570.000 a 574.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   47
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "520.000 a 539.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   46
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "SURFACTAN  VI"
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
         TabIndex        =   45
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Numeros de Lotes Reservados para cada Planta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   240
         Width           =   8295
      End
      Begin VB.Label Label27 
         Caption         =   "VERDE (OK)"
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
         Left            =   2520
         TabIndex        =   43
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "APROBADO"
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
         Left            =   2520
         TabIndex        =   42
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "AMARILLO (CR)"
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
         Left            =   4680
         TabIndex        =   41
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "Aprob. por Desvio"
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
         Left            =   4680
         TabIndex        =   40
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "ROJO (NO OK)"
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
         Left            =   6840
         TabIndex        =   39
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label32 
         Caption         =   "RECHAZADO"
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
         Left            =   6840
         TabIndex        =   38
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label33 
         Caption         =   "SURFACTAN  I"
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
         TabIndex        =   37
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label34 
         Caption         =   "SURFACTAN  I (Colorantes)"
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
         TabIndex        =   36
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label35 
         Caption         =   "SURFACTAN  II"
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
         TabIndex        =   35
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "SURFACTAN  III"
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
         TabIndex        =   34
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label37 
         Caption         =   "SURFACTAN  IV"
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
         TabIndex        =   33
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label38 
         Caption         =   "SURFACTAN  V"
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
         TabIndex        =   32
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label39 
         Caption         =   "PELLITAL I"
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
         TabIndex        =   31
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label40 
         Caption         =   "PELLITAL II"
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
         TabIndex        =   30
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label41 
         Caption         =   "PELLITAL III"
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
         TabIndex        =   29
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label42 
         Caption         =   "100.000 a 189.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   28
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label43 
         Caption         =   "900.000 a 989.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   27
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label44 
         Caption         =   "200.000 a 289.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   26
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label45 
         Caption         =   "300.000 a 389.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   25
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label46 
         Caption         =   "400.000 a 489.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   24
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label47 
         Caption         =   "800.000 a 889.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   23
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label48 
         Caption         =   "700.000 a 789.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label49 
         Caption         =   "600.000 a 689.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   21
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label50 
         Caption         =   "500.000 a 519.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label51 
         Caption         =   "390.000 a 394.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   19
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label52 
         Caption         =   "290.000 a 294.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label53 
         Caption         =   "990.000 a 994.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   17
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label54 
         Caption         =   "190.000 a 194.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   16
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label55 
         Caption         =   "490.000 a 494.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   15
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label56 
         Caption         =   "890.000 a 894.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label57 
         Caption         =   "790.000 a 794.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label58 
         Caption         =   "690.000 a 694.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label59 
         Caption         =   "590.000 a 594.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label60 
         Caption         =   "78.000   a 78.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label Label61 
         Caption         =   "74.000   a 74.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   9
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label62 
         Caption         =   "995.000 a 999.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label63 
         Caption         =   "70.000   a 70.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label64 
         Caption         =   "75.000   a 75.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label65 
         Caption         =   "73.000   a 73.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   5
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Label Label66 
         Caption         =   "76.000   a 76.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   4
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label Label67 
         Caption         =   "71.000   a 71.999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   3
         Top             =   4200
         Width           =   1695
      End
      Begin VB.Label Label68 
         Caption         =   "72.000   a 72.399"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   3000
         Width           =   1695
      End
   End
End
Attribute VB_Name = "PrgPrueArtRango"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FinNroLote_Click()
    PrgPrueArtRango.Hide
    Unload Me
    PrgPruart.Show
End Sub

