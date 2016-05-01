VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPruartrabajo 
   Caption         =   "Ingreso de Ensayos de Materia Prima"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8520
   ScaleWidth      =   11880
   Begin VB.TextBox Informe 
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
      Left            =   5520
      MaxLength       =   6
      TabIndex        =   207
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Desvio 
      Caption         =   "Traspaso a Desvio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   459
      Left            =   10320
      TabIndex        =   206
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Frame ImpreEtiqueta 
      Height          =   2175
      Left            =   600
      TabIndex        =   101
      Top             =   4200
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton AceptaEtiqueta 
         Caption         =   "Aceptar"
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
         Left            =   1320
         TabIndex        =   107
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton CancelaEtiqueta 
         Caption         =   "Cancelar"
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
         Left            =   3000
         TabIndex        =   106
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Kilos 
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
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   103
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Cantidad 
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
         Left            =   3120
         MaxLength       =   6
         TabIndex        =   102
         Text            =   " "
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label75 
         Caption         =   "Kilos Envase"
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
         Left            =   600
         TabIndex        =   105
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Cantidad de Etiquetas"
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
         Left            =   600
         TabIndex        =   104
         Top             =   720
         Width           =   1935
      End
   End
   Begin VB.Frame panLote 
      Caption         =   "Grabacion de Lote"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      TabIndex        =   26
      Top             =   4560
      Visible         =   0   'False
      Width           =   11535
      Begin VB.CommandButton btnConsultaPartida 
         Caption         =   "Consulta Partidas"
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
         Left            =   7800
         TabIndex        =   209
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox OrigenMercaderia 
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
         Left            =   8760
         MaxLength       =   30
         TabIndex        =   74
         Text            =   " "
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox PartidaProveedor 
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
         Left            =   6840
         MaxLength       =   20
         TabIndex        =   72
         Text            =   " "
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox NroRechazo 
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
         Left            =   4200
         MaxLength       =   6
         TabIndex        =   49
         Text            =   " "
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton CancelaLote 
         Caption         =   "Cancela Operacion"
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
         Left            =   5040
         TabIndex        =   36
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton GrabaLote 
         Caption         =   "Graba Prueba"
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
         Left            =   2160
         TabIndex        =   35
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Nueva 
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
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   34
         Text            =   " "
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox Devuelta 
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
         Left            =   2760
         MaxLength       =   10
         TabIndex        =   33
         Text            =   " "
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Liberada 
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   32
         Text            =   " "
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Lote 
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
         Left            =   120
         MaxLength       =   8
         TabIndex        =   31
         Text            =   " "
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "Origen Mercaderia"
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
         Left            =   9240
         TabIndex        =   75
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label24 
         Caption         =   "Nro Partida Proveedor"
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
         Left            =   6720
         TabIndex        =   73
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Nro Rechazo"
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
         Left            =   4200
         TabIndex        =   48
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "Nueva O/C"
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
         Left            =   5520
         TabIndex        =   30
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Canti.Devuelta"
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
         Left            =   2640
         TabIndex        =   29
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Cant.Liberada"
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
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Prueba"
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
         TabIndex        =   27
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Pass 
      Height          =   1815
      Left            =   840
      TabIndex        =   67
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
      Begin VB.CommandButton WCancela 
         Caption         =   "Cancela Grabacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   69
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox WClave 
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
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   68
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   70
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control de Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   720
      TabIndex        =   12
      Top             =   4440
      Visible         =   0   'False
      Width           =   5895
      Begin MSMask.MaskEdBox Hastafec 
         Height          =   300
         Left            =   4440
         TabIndex        =   56
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desdefec 
         Height          =   300
         Left            =   4440
         TabIndex        =   55
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   300
         Left            =   1560
         TabIndex        =   52
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.Frame Frame3 
         Caption         =   "Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   44
         Top             =   1200
         Width           =   1695
         Begin VB.OptionButton ImprePantalla 
            Caption         =   "Pantalla"
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
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton ImpreListado 
            Caption         =   "Listado"
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
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2160
         TabIndex        =   41
         Top             =   1200
         Width           =   1815
         Begin VB.OptionButton Rechazo 
            Caption         =   "Rechazados"
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
            TabIndex        =   43
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Aprobado 
            Caption         =   "Aprobados"
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
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   15
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label19 
         Caption         =   "Hasta fecha"
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
         TabIndex        =   54
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Desde Fecha"
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
         TabIndex        =   53
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Modif 
      Caption         =   "Modificacion de Orden de Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1200
      TabIndex        =   58
      Top             =   4440
      Visible         =   0   'False
      Width           =   4095
      Begin MSMask.MaskEdBox Modif_Recibido 
         Height          =   285
         Left            =   2400
         TabIndex        =   66
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Modif_Solicitado 
         Height          =   285
         Left            =   2400
         TabIndex        =   65
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.TextBox Modif_Orden 
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
         Left            =   2760
         MaxLength       =   6
         TabIndex        =   64
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton Modif_Cancela 
         Caption         =   "Cancela"
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
         Left            =   2280
         TabIndex        =   63
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Modif_Confirma 
         Caption         =   "Confirma "
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
         Left            =   480
         TabIndex        =   62
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Producto recibido"
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
         Left            =   240
         TabIndex        =   61
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Producto Solicitatado"
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
         Left            =   240
         TabIndex        =   60
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label20 
         Caption         =   "Orden de Compra"
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
         Left            =   240
         TabIndex        =   59
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.TextBox NroRevalida 
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
      Left            =   6120
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   111
      Text            =   " "
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Revalida 
      Caption         =   "Revalida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   459
      Left            =   10320
      TabIndex        =   110
      Top             =   5400
      Width           =   1455
   End
   Begin Crystal.CrystalReport ListaII 
      Left            =   11040
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Frame IngresaEstado 
      Caption         =   " "
      Height          =   3480
      Left            =   4440
      TabIndex        =   77
      Top             =   4560
      Visible         =   0   'False
      Width           =   5970
      Begin VB.CheckBox EstadoNo 
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
         Left            =   2760
         TabIndex        =   100
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox CertificadoNo 
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
         Left            =   2760
         TabIndex        =   99
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton ConfirmaEstado 
         Caption         =   "Confirma "
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
         Left            =   2640
         TabIndex        =   84
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox CertificadoSi 
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
         Left            =   2040
         TabIndex        =   81
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Certificado2 
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   80
         Text            =   " "
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox EstadoSi 
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
         Left            =   2040
         TabIndex        =   79
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox Estado2 
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
         Left            =   3480
         MaxLength       =   50
         TabIndex        =   78
         Text            =   " "
         Top             =   1680
         Width           =   2655
      End
      Begin MSMask.MaskEdBox Vencimiento 
         Height          =   285
         Left            =   1920
         TabIndex        =   108
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label76 
         Caption         =   "Vencimiento"
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
         Left            =   240
         TabIndex        =   109
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label71 
         Alignment       =   2  'Center
         Caption         =   "ACTUALIZACION DE DATOS DE CERTIFICADO DE ANALISIS Y ESTADO DE ENVASES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   85
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label70 
         Caption         =   "Certif.de Analisis"
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
         Left            =   240
         TabIndex        =   83
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label69 
         Caption         =   "Estado Envases"
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
         Left            =   240
         TabIndex        =   82
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.TextBox OrigenMercaderiaII 
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
      Left            =   9840
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   97
      Text            =   " "
      Top             =   360
      Width           =   1935
   End
   Begin VB.TextBox PartidaProveedorII 
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
      Left            =   10320
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   95
      Text            =   " "
      Top             =   0
      Width           =   1455
   End
   Begin VB.Frame PantaNumeroPrueba 
      Height          =   855
      Left            =   240
      TabIndex        =   92
      Top             =   4200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox NumeroPrueba 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   93
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label72 
         Caption         =   "Numero de Prueba"
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
         Left            =   240
         TabIndex        =   94
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.ListBox WPantalla 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      ItemData        =   "pruearttrabajo.frx":0000
      Left            =   7440
      List            =   "pruearttrabajo.frx":0007
      TabIndex        =   90
      Top             =   7800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   4
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   3
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   1
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox WTitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
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
      Index           =   2
      Left            =   10080
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   7680
      Visible         =   0   'False
      Width           =   375
   End
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
      Left            =   240
      TabIndex        =   76
      Top             =   5400
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Modifica 
      Caption         =   "Modificacion de  Prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   459
      Left            =   10320
      TabIndex        =   71
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Cambio 
      Caption         =   "  Modificacion          de O/C"
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
      Left            =   7200
      TabIndex        =   57
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Partida 
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
      Left            =   4680
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   51
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Impensayo 
      Caption         =   "Impresion Prueba"
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
      Left            =   8760
      TabIndex        =   47
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox Orden 
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
      Left            =   7320
      MaxLength       =   6
      TabIndex        =   40
      Text            =   " "
      Top             =   0
      Width           =   855
   End
   Begin MSMask.MaskEdBox fecha 
      Height          =   285
      Left            =   3000
      TabIndex        =   38
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdAddRechazo 
      Caption         =   "Graba Rechazo"
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
      Left            =   5640
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Confecciono 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   24
      Text            =   " "
      Top             =   4920
      Width           =   3975
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   23
      Text            =   " "
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox Aspecto 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   22
      Text            =   " "
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox Ensayo 
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
      Left            =   1560
      MaxLength       =   50
      TabIndex        =   21
      Text            =   " "
      Top             =   4200
      Width           =   3975
   End
   Begin MSMask.MaskEdBox Producto 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   10560
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wpruart.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   240
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   9600
      TabIndex        =   10
      Top             =   5640
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      ItemData        =   "pruearttrabajo.frx":0015
      Left            =   240
      List            =   "pruearttrabajo.frx":001C
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
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
      Left            =   10320
      TabIndex        =   5
      Top             =   4200
      Width           =   1455
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
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar Pantalla"
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
      Left            =   10320
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
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
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddlote 
      Caption         =   "Graba  Prueba"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   4800
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   735
      Left            =   8880
      TabIndex        =   91
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      _Version        =   327680
      BackColor       =   16777215
      ForeColor       =   4210752
      FocusRect       =   2
      GridLines       =   0
   End
   Begin MSMask.MaskEdBox Vto 
      Height          =   285
      Left            =   7320
      TabIndex        =   112
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   327680
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.TextBox RevalidaAnterior 
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
      Left            =   8640
      MaxLength       =   10
      TabIndex        =   115
      Text            =   " "
      Top             =   5400
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3300
      Left            =   120
      TabIndex        =   117
      Top             =   720
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5821
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificaciones 1 - 10"
      TabPicture(0)   =   "pruearttrabajo.frx":002A
      Tab(0).ControlCount=   44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Descri1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Descri2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Descri3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Descri4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Descri5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Descri6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Descri7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Descri8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Descri9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Descri10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label90"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Std1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Std2"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Std3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Std4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Std5"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Std6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Std7"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Std8"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Std9"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Std10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblresultado"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label92"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label26"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Valor1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Valor2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Valor3"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Valor4"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Valor5"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Valor6"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Valor7"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Valor8"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Valor9"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Valor10"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "ValorNumero1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "ValorNumero2"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ValorNumero3"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "ValorNumero4"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "ValorNumero5"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "ValorNumero6"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "ValorNumero7"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "ValorNumero8"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "ValorNumero9"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "ValorNumero10"
      Tab(0).Control(43).Enabled=   0   'False
      TabCaption(1)   =   "Especificaciones 11 - 20"
      TabPicture(1)   =   "pruearttrabajo.frx":0046
      Tab(1).ControlCount=   44
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ValorNumero20"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "ValorNumero19"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "ValorNumero18"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "ValorNumero17"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "ValorNumero16"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "ValorNumero15"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "ValorNumero14"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "ValorNumero13"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "ValorNumero12"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "ValorNumero11"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Valor20"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Valor19"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Valor18"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Valor17"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Valor16"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Valor15"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Valor14"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Valor13"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Valor12"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Valor11"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Label27"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label116"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "lblresultadoII"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Std20"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Std19"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Std18"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Std17"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Std16"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Std15"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Std14"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Std13"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Std12"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Std11"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label104"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Descri20"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Descri19"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "Descri18"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "Descri17"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Descri16"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Descri15"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).Control(40)=   "Descri14"
      Tab(1).Control(40).Enabled=   0   'False
      Tab(1).Control(41)=   "Descri13"
      Tab(1).Control(41).Enabled=   0   'False
      Tab(1).Control(42)=   "Descri12"
      Tab(1).Control(42).Enabled=   0   'False
      Tab(1).Control(43)=   "Descri11"
      Tab(1).Control(43).Enabled=   0   'False
      TabCaption(2)   =   "Especificaciones 21 - 30"
      TabPicture(2)   =   "pruearttrabajo.frx":0062
      Tab(2).ControlCount=   44
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label30"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Std30"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Std29"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Std28"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Std27"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Std26"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Std25"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Std24"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Std23"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Std22"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Std21"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label41"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Descri30"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Descri29"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Descri28"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Descri27"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Descri26"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Descri25"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Descri24"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Descri23"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Descri22"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Descri21"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "ValorNumero30"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "ValorNumero29"
      Tab(2).Control(25).Enabled=   -1  'True
      Tab(2).Control(26)=   "ValorNumero28"
      Tab(2).Control(26).Enabled=   -1  'True
      Tab(2).Control(27)=   "ValorNumero27"
      Tab(2).Control(27).Enabled=   -1  'True
      Tab(2).Control(28)=   "ValorNumero26"
      Tab(2).Control(28).Enabled=   -1  'True
      Tab(2).Control(29)=   "ValorNumero25"
      Tab(2).Control(29).Enabled=   -1  'True
      Tab(2).Control(30)=   "ValorNumero24"
      Tab(2).Control(30).Enabled=   -1  'True
      Tab(2).Control(31)=   "ValorNumero23"
      Tab(2).Control(31).Enabled=   -1  'True
      Tab(2).Control(32)=   "ValorNumero22"
      Tab(2).Control(32).Enabled=   -1  'True
      Tab(2).Control(33)=   "ValorNumero21"
      Tab(2).Control(33).Enabled=   -1  'True
      Tab(2).Control(34)=   "Valor30"
      Tab(2).Control(34).Enabled=   -1  'True
      Tab(2).Control(35)=   "Valor29"
      Tab(2).Control(35).Enabled=   -1  'True
      Tab(2).Control(36)=   "Valor28"
      Tab(2).Control(36).Enabled=   -1  'True
      Tab(2).Control(37)=   "Valor27"
      Tab(2).Control(37).Enabled=   -1  'True
      Tab(2).Control(38)=   "Valor26"
      Tab(2).Control(38).Enabled=   -1  'True
      Tab(2).Control(39)=   "Valor25"
      Tab(2).Control(39).Enabled=   -1  'True
      Tab(2).Control(40)=   "Valor24"
      Tab(2).Control(40).Enabled=   -1  'True
      Tab(2).Control(41)=   "Valor23"
      Tab(2).Control(41).Enabled=   -1  'True
      Tab(2).Control(42)=   "Valor22"
      Tab(2).Control(42).Enabled=   -1  'True
      Tab(2).Control(43)=   "Valor21"
      Tab(2).Control(43).Enabled=   -1  'True
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   229
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   228
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   227
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   226
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   225
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   224
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   223
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   222
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   221
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   220
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox ValorNumero21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   219
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox ValorNumero22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   218
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox ValorNumero23 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   217
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox ValorNumero24 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   216
         Top             =   1440
         Width           =   800
      End
      Begin VB.TextBox ValorNumero25 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   215
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox ValorNumero26 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   214
         Top             =   1920
         Width           =   800
      End
      Begin VB.TextBox ValorNumero27 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   213
         Top             =   2160
         Width           =   800
      End
      Begin VB.TextBox ValorNumero28 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   212
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox ValorNumero29 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   211
         Top             =   2640
         Width           =   800
      End
      Begin VB.TextBox ValorNumero30 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -64245
         MaxLength       =   8
         TabIndex        =   210
         Top             =   2880
         Width           =   800
      End
      Begin VB.TextBox ValorNumero20 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   204
         Top             =   2880
         Width           =   800
      End
      Begin VB.TextBox ValorNumero19 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   203
         Top             =   2640
         Width           =   800
      End
      Begin VB.TextBox ValorNumero18 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   202
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox ValorNumero17 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   201
         Top             =   2160
         Width           =   800
      End
      Begin VB.TextBox ValorNumero16 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   200
         Top             =   1920
         Width           =   800
      End
      Begin VB.TextBox ValorNumero15 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   199
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox ValorNumero14 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   198
         Top             =   1440
         Width           =   800
      End
      Begin VB.TextBox ValorNumero13 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   197
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox ValorNumero12 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   196
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox ValorNumero11 
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
         Left            =   -64250
         MaxLength       =   8
         TabIndex        =   195
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox ValorNumero10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   193
         Top             =   2880
         Width           =   800
      End
      Begin VB.TextBox ValorNumero9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   192
         Top             =   2640
         Width           =   800
      End
      Begin VB.TextBox ValorNumero8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   191
         Top             =   2400
         Width           =   800
      End
      Begin VB.TextBox ValorNumero7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   190
         Top             =   2160
         Width           =   800
      End
      Begin VB.TextBox ValorNumero6 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   189
         Top             =   1920
         Width           =   800
      End
      Begin VB.TextBox ValorNumero5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   188
         Top             =   1680
         Width           =   800
      End
      Begin VB.TextBox ValorNumero4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   187
         Top             =   1440
         Width           =   800
      End
      Begin VB.TextBox ValorNumero3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   186
         Top             =   1200
         Width           =   800
      End
      Begin VB.TextBox ValorNumero2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   185
         Top             =   960
         Width           =   800
      End
      Begin VB.TextBox ValorNumero1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10750
         MaxLength       =   8
         TabIndex        =   184
         Top             =   720
         Width           =   800
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   183
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   182
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   181
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   180
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   179
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   178
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   177
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   176
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   175
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -65940
         Locked          =   -1  'True
         TabIndex        =   174
         Top             =   720
         Width           =   1700
      End
      Begin VB.TextBox Valor10 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   150
         Top             =   2880
         Width           =   1700
      End
      Begin VB.TextBox Valor9 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   149
         Top             =   2640
         Width           =   1700
      End
      Begin VB.TextBox Valor8 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   148
         Top             =   2400
         Width           =   1700
      End
      Begin VB.TextBox Valor7 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   147
         Top             =   2160
         Width           =   1700
      End
      Begin VB.TextBox Valor6 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   146
         Top             =   1920
         Width           =   1700
      End
      Begin VB.TextBox Valor5 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   145
         Top             =   1680
         Width           =   1700
      End
      Begin VB.TextBox Valor4 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   144
         Top             =   1440
         Width           =   1700
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   143
         Top             =   1200
         Width           =   1700
      End
      Begin VB.TextBox Valor2 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   142
         Top             =   960
         Width           =   1700
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   9060
         Locked          =   -1  'True
         TabIndex        =   141
         Top             =   720
         Width           =   1700
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   253
         Top             =   720
         Width           =   3780
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   252
         Top             =   960
         Width           =   3780
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   251
         Top             =   1200
         Width           =   3780
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   250
         Top             =   1440
         Width           =   3780
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   249
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   248
         Top             =   1920
         Width           =   3780
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   247
         Top             =   2160
         Width           =   3780
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   246
         Top             =   2400
         Width           =   3780
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   245
         Top             =   2640
         Width           =   3780
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   244
         Top             =   2880
         Width           =   3780
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   243
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label Std21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   242
         Top             =   720
         Width           =   5040
      End
      Begin VB.Label Std22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   241
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Std23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   240
         Top             =   1200
         Width           =   5040
      End
      Begin VB.Label Std24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   239
         Top             =   1440
         Width           =   5040
      End
      Begin VB.Label Std25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   238
         Top             =   1680
         Width           =   5040
      End
      Begin VB.Label Std26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   237
         Top             =   1920
         Width           =   5040
      End
      Begin VB.Label Std27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   236
         Top             =   2160
         Width           =   5040
      End
      Begin VB.Label Std28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   235
         Top             =   2400
         Width           =   5040
      End
      Begin VB.Label Std29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   234
         Top             =   2640
         Width           =   5040
      End
      Begin VB.Label Std30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -70995
         TabIndex        =   233
         Top             =   2880
         Width           =   5040
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -70995
         TabIndex        =   232
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   -65940
         TabIndex        =   231
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -64245
         TabIndex        =   230
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -64250
         TabIndex        =   205
         Top             =   480
         Width           =   795
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   10750
         TabIndex        =   194
         Top             =   480
         Width           =   800
      End
      Begin VB.Label Label116 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   -65940
         TabIndex        =   173
         Top             =   480
         Width           =   1700
      End
      Begin VB.Label lblresultadoII 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   -71000
         TabIndex        =   172
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Std20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   171
         Top             =   2880
         Width           =   5040
      End
      Begin VB.Label Std19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   170
         Top             =   2640
         Width           =   5040
      End
      Begin VB.Label Std18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   169
         Top             =   2400
         Width           =   5040
      End
      Begin VB.Label Std17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   168
         Top             =   2160
         Width           =   5040
      End
      Begin VB.Label Std16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   167
         Top             =   1920
         Width           =   5040
      End
      Begin VB.Label Std15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   166
         Top             =   1680
         Width           =   5040
      End
      Begin VB.Label Std14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   165
         Top             =   1440
         Width           =   5040
      End
      Begin VB.Label Std13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   164
         Top             =   1200
         Width           =   5040
      End
      Begin VB.Label Std12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   163
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Std11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -71000
         TabIndex        =   162
         Top             =   720
         Width           =   5040
      End
      Begin VB.Label Label104 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   161
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   160
         Top             =   2880
         Width           =   3780
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   159
         Top             =   2640
         Width           =   3780
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   158
         Top             =   2400
         Width           =   3780
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   157
         Top             =   2160
         Width           =   3780
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   156
         Top             =   1920
         Width           =   3780
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   155
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   154
         Top             =   1440
         Width           =   3780
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   153
         Top             =   1200
         Width           =   3780
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   152
         Top             =   960
         Width           =   3780
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   -74880
         TabIndex        =   151
         Top             =   720
         Width           =   3780
      End
      Begin VB.Label Label92 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Registrado"
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
         Left            =   9060
         TabIndex        =   140
         Top             =   480
         Width           =   1700
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   4000
         TabIndex        =   139
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Std10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   138
         Top             =   2880
         Width           =   5040
      End
      Begin VB.Label Std9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   137
         Top             =   2640
         Width           =   5040
      End
      Begin VB.Label Std8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4005
         TabIndex        =   136
         Top             =   2400
         Width           =   5040
      End
      Begin VB.Label Std7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   135
         Top             =   2160
         Width           =   5040
      End
      Begin VB.Label Std6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   134
         Top             =   1920
         Width           =   5040
      End
      Begin VB.Label Std5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   133
         Top             =   1680
         Width           =   5040
      End
      Begin VB.Label Std4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   132
         Top             =   1440
         Width           =   5040
      End
      Begin VB.Label Std3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   131
         Top             =   1200
         Width           =   5040
      End
      Begin VB.Label Std2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   130
         Top             =   960
         Width           =   5040
      End
      Begin VB.Label Std1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   4000
         TabIndex        =   129
         Top             =   720
         Width           =   5040
      End
      Begin VB.Label Label90 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         TabIndex        =   128
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   127
         Top             =   2880
         Width           =   3780
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   2640
         Width           =   3780
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   2400
         Width           =   3780
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   124
         Top             =   2160
         Width           =   3780
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   123
         Top             =   1920
         Width           =   3780
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   1680
         Width           =   3780
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   1440
         Width           =   3780
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   120
         Top             =   1200
         Width           =   3780
      End
      Begin VB.Label Descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   960
         Width           =   3780
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   720
         Width           =   3780
      End
   End
   Begin VB.Label Label28 
      Caption         =   "Informe"
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
      TabIndex        =   208
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label78 
      Caption         =   "Vto."
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
      Left            =   6960
      TabIndex        =   114
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label77 
      Caption         =   "Rev."
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
      Left            =   5760
      TabIndex        =   113
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label74 
      Caption         =   "Origen Merc."
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
      Left            =   8760
      TabIndex        =   98
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label73 
      Caption         =   "Nro Partida Proveedor"
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
      Left            =   8280
      TabIndex        =   96
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Lote"
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
      Left            =   4080
      TabIndex        =   50
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Orden"
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
      Left            =   6480
      TabIndex        =   39
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha"
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
      Left            =   2280
      TabIndex        =   37
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Confecciono"
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
      TabIndex        =   20
      Top             =   4920
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Label Label5 
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
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label Label79 
      Caption         =   "Rev."
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
      Left            =   7920
      TabIndex        =   116
      Top             =   5400
      Width           =   615
   End
End
Attribute VB_Name = "PrgPruartrabajo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private WInforme As String
Private Pasa As String
Private Auxi3 As String
Private Auxi4 As String
Private WLote As String
Private ZLote As String
Private SaldoOrden As Double
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim rstPrueart As Recordset
Dim spPrueart As String
Dim rstLaudo As Recordset
Dim spLaudo As String
Dim rstOrden As Recordset
Dim spOrden As String
Dim rstInforme As Recordset
Dim spInforme As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecificacionesUnifica As Recordset
Dim spEspecificacionesUnifica As String
Dim rstRevalida As Recordset
Dim spRevalida As String
Dim XParam As String
Dim WCosto1 As String
Dim WCosto3 As String
Dim WPrecio As Double
Dim XStock As Double
Dim XCosto As Double
Dim XCostoTotal As Double
Dim XStock1 As Double
Dim XCosto1 As Double
Dim XCostoTotal1 As Double
Dim XStock2 As Double
Dim XCosto2 As Double
Dim XCostoTotal2 As Double
Dim XCosto3 As Double
Dim WTipoOrden As Single
Dim WOrigen  As String
Dim WRecibida As Double
Dim WLaudada As Double
Dim WCantidad1 As Double
Dim WCantidad2 As Double
Dim XCertificado1 As Integer
Dim XCertificado2 As String
Dim XEstado1 As Integer
Dim XEstado2 As String
Dim WProceso As Integer
Dim ZVencimiento As String
Dim XMes As String
Dim XAno As String
Dim ZCantidad As String
Dim ZNroDespacho As String
Dim ZProcedencia As String

Dim ZEnsayo1 As String
Dim ZEnsayo2 As String
Dim ZEnsayo3 As String
Dim ZEnsayo4 As String
Dim ZEnsayo5 As String
Dim ZEnsayo6 As String
Dim ZEnsayo7 As String
Dim ZEnsayo8 As String
Dim ZEnsayo9 As String
Dim ZEnsayo10 As String
Dim ZEnsayo11 As String
Dim ZEnsayo12 As String
Dim ZEnsayo13 As String
Dim ZEnsayo14 As String
Dim ZEnsayo15 As String
Dim ZEnsayo16 As String
Dim ZEnsayo17 As String
Dim ZEnsayo18 As String
Dim ZEnsayo19 As String
Dim ZEnsayo20 As String
Dim ZEnsayo21 As String
Dim ZEnsayo22 As String
Dim ZEnsayo23 As String
Dim ZEnsayo24 As String
Dim ZEnsayo25 As String
Dim ZEnsayo26 As String
Dim ZEnsayo27 As String
Dim ZEnsayo28 As String
Dim ZEnsayo29 As String
Dim ZEnsayo30 As String



Dim ZEnsayo(30) As String
Dim ZDesde(30) As String
Dim ZHasta(30) As String
Dim ZUnidad(30) As String
Dim ZValorNumero(30) As String

Dim EmpresaActual As String

Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim WDireccionEmail As String
Dim EmailAddress As String
Dim CopiaAddress As String
Dim WNombreEmail As String
Dim MAttach As String

Private Sub Acepta_Click()

    Desde.Text = UCase(Desde.Text)
    
    If Aprobado.Value = True Then
        Desdepru = "100000"
        HastaPru = "199999"
            Else
        Desdepru = "200000"
        HastaPru = "299999"
    End If
    
    WAno = Right$(Desdefec.Text, 4)
    WMes = Mid$(Desdefec.Text, 4, 2)
    WDia = Left$(Desdefec.Text, 2)
    FDesde = WAno + WMes + WDia
    WAno = Right$(Hastafec.Text, 4)
    WMes = Mid$(Hastafec.Text, 4, 2)
    WDia = Left$(Hastafec.Text, 2)
    FHasta = WAno + WMes + WDia
    
    With rstPrueba
        .Index = "Clave"
        .Seek ">=", "0"
        If .NoMatch = False Then
            Do
                .Delete
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Erase ZDesde
    Erase ZHasta

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Desde.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        ZEnsayo1 = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo2 = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo3 = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo4 = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo5 = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo6 = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo7 = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo8 = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo9 = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo10 = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        ZDesde(1) = IIf(IsNull(rstEspecificacionesUnifica!Desde1), "", rstEspecificacionesUnifica!Desde1)
        ZDesde(2) = IIf(IsNull(rstEspecificacionesUnifica!Desde2), "", rstEspecificacionesUnifica!Desde2)
        ZDesde(3) = IIf(IsNull(rstEspecificacionesUnifica!Desde3), "", rstEspecificacionesUnifica!Desde3)
        ZDesde(4) = IIf(IsNull(rstEspecificacionesUnifica!Desde4), "", rstEspecificacionesUnifica!Desde4)
        ZDesde(5) = IIf(IsNull(rstEspecificacionesUnifica!Desde5), "", rstEspecificacionesUnifica!Desde5)
        ZDesde(6) = IIf(IsNull(rstEspecificacionesUnifica!Desde6), "", rstEspecificacionesUnifica!Desde6)
        ZDesde(7) = IIf(IsNull(rstEspecificacionesUnifica!Desde7), "", rstEspecificacionesUnifica!Desde7)
        ZDesde(8) = IIf(IsNull(rstEspecificacionesUnifica!Desde8), "", rstEspecificacionesUnifica!Desde8)
        ZDesde(9) = IIf(IsNull(rstEspecificacionesUnifica!Desde9), "", rstEspecificacionesUnifica!Desde9)
        ZDesde(10) = IIf(IsNull(rstEspecificacionesUnifica!Desde10), "", rstEspecificacionesUnifica!Desde10)
        ZDesde(11) = IIf(IsNull(rstEspecificacionesUnifica!Desde11), "", rstEspecificacionesUnifica!Desde11)
        ZDesde(12) = IIf(IsNull(rstEspecificacionesUnifica!Desde12), "", rstEspecificacionesUnifica!Desde12)
        ZDesde(13) = IIf(IsNull(rstEspecificacionesUnifica!Desde13), "", rstEspecificacionesUnifica!Desde13)
        ZDesde(14) = IIf(IsNull(rstEspecificacionesUnifica!Desde14), "", rstEspecificacionesUnifica!Desde14)
        ZDesde(15) = IIf(IsNull(rstEspecificacionesUnifica!Desde15), "", rstEspecificacionesUnifica!Desde15)
        ZDesde(16) = IIf(IsNull(rstEspecificacionesUnifica!Desde16), "", rstEspecificacionesUnifica!Desde16)
        ZDesde(17) = IIf(IsNull(rstEspecificacionesUnifica!Desde17), "", rstEspecificacionesUnifica!Desde17)
        ZDesde(18) = IIf(IsNull(rstEspecificacionesUnifica!Desde18), "", rstEspecificacionesUnifica!Desde18)
        ZDesde(19) = IIf(IsNull(rstEspecificacionesUnifica!Desde19), "", rstEspecificacionesUnifica!Desde19)
        ZDesde(20) = IIf(IsNull(rstEspecificacionesUnifica!Desde20), "", rstEspecificacionesUnifica!Desde20)
        
        ZHasta(1) = IIf(IsNull(rstEspecificacionesUnifica!Hasta1), "", rstEspecificacionesUnifica!Hasta1)
        ZHasta(2) = IIf(IsNull(rstEspecificacionesUnifica!Hasta2), "", rstEspecificacionesUnifica!Hasta2)
        ZHasta(3) = IIf(IsNull(rstEspecificacionesUnifica!Hasta3), "", rstEspecificacionesUnifica!Hasta3)
        ZHasta(4) = IIf(IsNull(rstEspecificacionesUnifica!Hasta4), "", rstEspecificacionesUnifica!Hasta4)
        ZHasta(5) = IIf(IsNull(rstEspecificacionesUnifica!Hasta5), "", rstEspecificacionesUnifica!Hasta5)
        ZHasta(6) = IIf(IsNull(rstEspecificacionesUnifica!Hasta6), "", rstEspecificacionesUnifica!Hasta6)
        ZHasta(7) = IIf(IsNull(rstEspecificacionesUnifica!Hasta7), "", rstEspecificacionesUnifica!Hasta7)
        ZHasta(8) = IIf(IsNull(rstEspecificacionesUnifica!Hasta8), "", rstEspecificacionesUnifica!Hasta8)
        ZHasta(9) = IIf(IsNull(rstEspecificacionesUnifica!Hasta9), "", rstEspecificacionesUnifica!Hasta9)
        ZHasta(10) = IIf(IsNull(rstEspecificacionesUnifica!Hasta10), "", rstEspecificacionesUnifica!Hasta10)
        ZHasta(11) = IIf(IsNull(rstEspecificacionesUnifica!Hasta11), "", rstEspecificacionesUnifica!Hasta11)
        ZHasta(12) = IIf(IsNull(rstEspecificacionesUnifica!Hasta12), "", rstEspecificacionesUnifica!Hasta12)
        ZHasta(13) = IIf(IsNull(rstEspecificacionesUnifica!Hasta13), "", rstEspecificacionesUnifica!Hasta13)
        ZHasta(14) = IIf(IsNull(rstEspecificacionesUnifica!Hasta14), "", rstEspecificacionesUnifica!Hasta14)
        ZHasta(15) = IIf(IsNull(rstEspecificacionesUnifica!Hasta15), "", rstEspecificacionesUnifica!Hasta15)
        ZHasta(16) = IIf(IsNull(rstEspecificacionesUnifica!Hasta16), "", rstEspecificacionesUnifica!Hasta16)
        ZHasta(17) = IIf(IsNull(rstEspecificacionesUnifica!Hasta17), "", rstEspecificacionesUnifica!Hasta17)
        ZHasta(18) = IIf(IsNull(rstEspecificacionesUnifica!Hasta18), "", rstEspecificacionesUnifica!Hasta18)
        ZHasta(19) = IIf(IsNull(rstEspecificacionesUnifica!Hasta19), "", rstEspecificacionesUnifica!Hasta19)
        ZHasta(20) = IIf(IsNull(rstEspecificacionesUnifica!Hasta20), "", rstEspecificacionesUnifica!Hasta20)
        
        ZDesde(1) = Trim(ZDesde(1))
        ZDesde(2) = Trim(ZDesde(2))
        ZDesde(3) = Trim(ZDesde(3))
        ZDesde(4) = Trim(ZDesde(4))
        ZDesde(5) = Trim(ZDesde(5))
        ZDesde(6) = Trim(ZDesde(6))
        ZDesde(7) = Trim(ZDesde(7))
        ZDesde(8) = Trim(ZDesde(8))
        ZDesde(9) = Trim(ZDesde(9))
        ZDesde(10) = Trim(ZDesde(10))
        ZDesde(11) = Trim(ZDesde(11))
        ZDesde(12) = Trim(ZDesde(12))
        ZDesde(13) = Trim(ZDesde(13))
        ZDesde(14) = Trim(ZDesde(14))
        ZDesde(15) = Trim(ZDesde(15))
        ZDesde(16) = Trim(ZDesde(16))
        ZDesde(17) = Trim(ZDesde(17))
        ZDesde(18) = Trim(ZDesde(18))
        ZDesde(19) = Trim(ZDesde(19))
        ZDesde(20) = Trim(ZDesde(20))
        
        ZHasta(1) = Trim(ZHasta(1))
        ZHasta(2) = Trim(ZHasta(2))
        ZHasta(3) = Trim(ZHasta(3))
        ZHasta(4) = Trim(ZHasta(4))
        ZHasta(5) = Trim(ZHasta(5))
        ZHasta(6) = Trim(ZHasta(6))
        ZHasta(7) = Trim(ZHasta(7))
        ZHasta(8) = Trim(ZHasta(8))
        ZHasta(9) = Trim(ZHasta(9))
        ZHasta(10) = Trim(ZHasta(10))
        ZHasta(11) = Trim(ZHasta(11))
        ZHasta(12) = Trim(ZHasta(12))
        ZHasta(13) = Trim(ZHasta(13))
        ZHasta(14) = Trim(ZHasta(14))
        ZHasta(15) = Trim(ZHasta(15))
        ZHasta(16) = Trim(ZHasta(16))
        ZHasta(17) = Trim(ZHasta(17))
        ZHasta(18) = Trim(ZHasta(18))
        ZHasta(19) = Trim(ZHasta(19))
        ZHasta(20) = Trim(ZHasta(20))
        
        
        rstEspecificacionesUnifica.Close
    End If
    
    
    
    
    

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Desde.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        ZEnsayo21 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo22 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo23 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo24 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo25 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo26 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo27 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo28 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo29 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo30 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        ZDesde(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde21), "", rstEspecificacionesUnificaIII!Desde21)
        ZDesde(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde22), "", rstEspecificacionesUnificaIII!Desde22)
        ZDesde(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde23), "", rstEspecificacionesUnificaIII!Desde23)
        ZDesde(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde24), "", rstEspecificacionesUnificaIII!Desde24)
        ZDesde(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde25), "", rstEspecificacionesUnificaIII!Desde25)
        ZDesde(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde26), "", rstEspecificacionesUnificaIII!Desde26)
        ZDesde(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde27), "", rstEspecificacionesUnificaIII!Desde27)
        ZDesde(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde28), "", rstEspecificacionesUnificaIII!Desde28)
        ZDesde(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde29), "", rstEspecificacionesUnificaIII!Desde29)
        ZDesde(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde30), "", rstEspecificacionesUnificaIII!Desde30)
        
        ZHasta(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta21), "", rstEspecificacionesUnificaIII!Hasta21)
        ZHasta(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta22), "", rstEspecificacionesUnificaIII!Hasta22)
        ZHasta(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta23), "", rstEspecificacionesUnificaIII!Hasta23)
        ZHasta(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta24), "", rstEspecificacionesUnificaIII!Hasta24)
        ZHasta(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta25), "", rstEspecificacionesUnificaIII!Hasta25)
        ZHasta(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta26), "", rstEspecificacionesUnificaIII!Hasta26)
        ZHasta(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta27), "", rstEspecificacionesUnificaIII!Hasta27)
        ZHasta(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta28), "", rstEspecificacionesUnificaIII!Hasta28)
        ZHasta(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta29), "", rstEspecificacionesUnificaIII!Hasta29)
        ZHasta(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta30), "", rstEspecificacionesUnificaIII!Hasta30)
        
        ZDesde(21) = Trim(ZDesde(21))
        ZDesde(22) = Trim(ZDesde(22))
        ZDesde(23) = Trim(ZDesde(23))
        ZDesde(24) = Trim(ZDesde(24))
        ZDesde(25) = Trim(ZDesde(25))
        ZDesde(26) = Trim(ZDesde(26))
        ZDesde(27) = Trim(ZDesde(27))
        ZDesde(28) = Trim(ZDesde(28))
        ZDesde(29) = Trim(ZDesde(29))
        ZDesde(30) = Trim(ZDesde(30))
        
        ZHasta(21) = Trim(ZHasta(21))
        ZHasta(22) = Trim(ZHasta(22))
        ZHasta(23) = Trim(ZHasta(23))
        ZHasta(24) = Trim(ZHasta(24))
        ZHasta(25) = Trim(ZHasta(25))
        ZHasta(26) = Trim(ZHasta(26))
        ZHasta(27) = Trim(ZHasta(27))
        ZHasta(28) = Trim(ZHasta(28))
        ZHasta(29) = Trim(ZHasta(29))
        ZHasta(30) = Trim(ZHasta(30))
        
        rstEspecificacionesUnificaIII.Close
    End If
    
    
    
    
    
    
    
    
    
    ZDesValor1 = ""
    ZDesValor2 = ""
    ZDesValor3 = ""
    ZDesValor4 = ""
    ZDesValor5 = ""
    ZDesValor6 = ""
    ZDesValor7 = ""
    ZDesValor8 = ""
    ZDesValor9 = ""
    ZDesValor10 = ""
    ZDesValor11 = ""
    ZDesValor12 = ""
    ZDesValor13 = ""
    ZDesValor14 = ""
    ZDesValor15 = ""
    ZDesValor16 = ""
    ZDesValor17 = ""
    ZDesValor18 = ""
    ZDesValor19 = ""
    ZDesValor20 = ""
    ZDesValor21 = ""
    ZDesValor22 = ""
    ZDesValor23 = ""
    ZDesValor24 = ""
    ZDesValor25 = ""
    ZDesValor26 = ""
    ZDesValor27 = ""
    ZDesValor28 = ""
    ZDesValor29 = ""
    ZDesValor30 = ""
    
    For Cicla = 1 To 30
        ZZDescri = ""
        If Val(ZEnsayo(Cicla)) <> 0 Then
            spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDescri = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        End If
        Select Case Cicla
            Case 1
                ZDesValor1 = ZZDescri
            Case 2
                ZDesValor2 = ZZDescri
            Case 3
                ZDesValor3 = ZZDescri
            Case 4
                ZDesValor4 = ZZDescri
            Case 5
                ZDesValor5 = ZZDescri
            Case 6
                ZDesValor6 = ZZDescri
            Case 7
                ZDesValor7 = ZZDescri
            Case 8
                ZDesValor8 = ZZDescri
            Case 9
                ZDesValor9 = ZZDescri
            Case 10
                ZDesValor10 = ZZDescri
            Case 11
                ZDesValor11 = ZZDescri
            Case 12
                ZDesValor12 = ZZDescri
            Case 13
                ZDesValor13 = ZZDescri
            Case 14
                ZDesValor14 = ZZDescri
            Case 15
                ZDesValor15 = ZZDescri
            Case 16
                ZDesValor16 = ZZDescri
            Case 17
                ZDesValor17 = ZZDescri
            Case 18
                ZDesValor18 = ZZDescri
            Case 19
                ZDesValor19 = ZZDescri
            Case 20
                ZDesValor20 = ZZDescri
            Case 21
                ZDesValor21 = ZZDescri
            Case 22
                ZDesValor22 = ZZDescri
            Case 23
                ZDesValor23 = ZZDescri
            Case 24
                ZDesValor24 = ZZDescri
            Case 25
                ZDesValor25 = ZZDescri
            Case 26
                ZDesValor26 = ZZDescri
            Case 27
                ZDesValor27 = ZZDescri
            Case 28
                ZDesValor28 = ZZDescri
            Case 29
                ZDesValor29 = ZZDescri
            Case 30
                ZDesValor30 = ZZDescri
            Case Else
        End Select
                
    Next Cicla
    
    Call Conecta_Empresa
    
    Suma = 0
    
    ZSql = ""
    ZSql = "Select Prueart.Prueba, Prueart.Producto, Prueart.Fecha, Prueart.Orden, Prueart.Valor1,  Prueart.Valor2,  Prueart.Valor3,  Prueart.Valor4,  Prueart.Valor5,  Prueart.Valor6,  Prueart.Valor7,  Prueart.Valor8,  Prueart.Valor9,  Prueart.Valor10,  Prueart.Valor11,  Prueart.Valor12,  Prueart.Valor13,  Prueart.Valor14,  Prueart.Valor15,  Prueart.Valor16,  Prueart.Valor17,  Prueart.Valor18,  Prueart.Valor19, Prueart.Valor20, Prueart.Valor21, Prueart.Valor22, Prueart.Valor23, Prueart.Valor24, Prueart.Valor25, Prueart.Valor26, Prueart.Valor27, Prueart.Valor28, Prueart.Valor29, Prueart.Valor30, Articulo.Descripcion as [DesProducto]  "
    ZSql = ZSql & " FROM Prueart, Articulo"
    ZSql = ZSql & " Where Prueart.Producto = " + "'" + Desde.Text + "'"
    ZSql = ZSql & " and Prueart.FechaOrd >= " + "'" + FDesde + "'"
    ZSql = ZSql & " and Prueart.FechaOrd <= " + "'" + FHasta + "'"
    ZSql = ZSql & " and Prueart.Producto = Articulo.Codigo"
    
    spPrueart = ZSql
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
    
        With rstPrueart
            .MoveFirst
            If .NoMatch = False Then
                Do
            
                    If .EOF = True Then
                        Exit Do
                    End If
                    
                    Suma = Suma + 1
                    
                    ZPrueba = rstPrueart!Prueba
                    ZProducto = rstPrueart!Producto
                    ZFecha = rstPrueart!fecha
                    ZOrden = rstPrueart!Orden
                    ZValor1 = rstPrueart!Valor1
                    ZValor2 = rstPrueart!Valor2
                    ZValor3 = rstPrueart!Valor3
                    ZValor4 = rstPrueart!Valor4
                    ZValor5 = rstPrueart!Valor5
                    ZValor6 = rstPrueart!Valor6
                    ZValor7 = rstPrueart!Valor7
                    ZValor8 = rstPrueart!Valor8
                    ZValor9 = rstPrueart!Valor9
                    ZValor10 = rstPrueart!Valor10
                    ZValor11 = IIf(IsNull(rstPrueart!Valor11), "", rstPrueart!Valor11)
                    ZValor12 = IIf(IsNull(rstPrueart!Valor12), "", rstPrueart!Valor12)
                    ZValor13 = IIf(IsNull(rstPrueart!Valor13), "", rstPrueart!Valor13)
                    ZValor14 = IIf(IsNull(rstPrueart!Valor14), "", rstPrueart!Valor14)
                    ZValor15 = IIf(IsNull(rstPrueart!Valor15), "", rstPrueart!Valor15)
                    ZValor16 = IIf(IsNull(rstPrueart!Valor16), "", rstPrueart!Valor16)
                    ZValor17 = IIf(IsNull(rstPrueart!Valor17), "", rstPrueart!Valor17)
                    ZValor18 = IIf(IsNull(rstPrueart!Valor18), "", rstPrueart!Valor18)
                    ZValor19 = IIf(IsNull(rstPrueart!Valor19), "", rstPrueart!Valor19)
                    ZValor20 = IIf(IsNull(rstPrueart!Valor20), "", rstPrueart!Valor20)
                    ZValor21 = IIf(IsNull(rstPrueart!Valor21), "", rstPrueart!Valor21)
                    ZValor22 = IIf(IsNull(rstPrueart!Valor22), "", rstPrueart!Valor22)
                    ZValor23 = IIf(IsNull(rstPrueart!Valor23), "", rstPrueart!Valor23)
                    ZValor24 = IIf(IsNull(rstPrueart!Valor24), "", rstPrueart!Valor24)
                    ZValor25 = IIf(IsNull(rstPrueart!Valor25), "", rstPrueart!Valor25)
                    ZValor26 = IIf(IsNull(rstPrueart!Valor26), "", rstPrueart!Valor26)
                    ZValor27 = IIf(IsNull(rstPrueart!Valor27), "", rstPrueart!Valor27)
                    ZValor28 = IIf(IsNull(rstPrueart!Valor28), "", rstPrueart!Valor28)
                    ZValor29 = IIf(IsNull(rstPrueart!Valor29), "", rstPrueart!Valor29)
                    ZValor30 = IIf(IsNull(rstPrueart!Valor30), "", rstPrueart!Valor30)
                    
                    ZDesProducto = rstPrueart!DesProducto
                    
                    With rstPrueba
                        .AddNew
                        !Clave = Suma
                        !Prueba = ZPrueba
                        !Producto = ZProducto
                        !fecha = ZFecha
                        !Orden = ZOrden
                        !Valor1 = ZValor1
                        !Valor2 = ZValor2
                        !Valor3 = ZValor3
                        !Valor4 = ZValor4
                        !Valor5 = ZValor5
                        !Valor6 = ZValor6
                        !Valor7 = ZValor7
                        !Valor8 = ZValor8
                        !Valor9 = ZValor9
                        !Valor10 = ZValor10
                        !Valor11 = ZValor11
                        !Valor12 = ZValor12
                        !Valor13 = ZValor13
                        !Valor14 = ZValor14
                        !Valor15 = ZValor15
                        !Valor16 = ZValor16
                        !Valor17 = ZValor17
                        !Valor18 = ZValor18
                        !Valor19 = ZValor19
                        !Valor20 = ZValor20
                        !Valor21 = ZValor21
                        !Valor22 = ZValor22
                        !Valor23 = ZValor23
                        !Valor24 = ZValor24
                        !Valor25 = ZValor25
                        !Valor26 = ZValor26
                        !Valor27 = ZValor27
                        !Valor28 = ZValor28
                        !Valor29 = ZValor29
                        !Valor30 = ZValor30
                        !DesValor1 = ZDesValor1
                        !Desvalor2 = ZDesValor2
                        !DesValor3 = ZDesValor3
                        !Desvalor4 = ZDesValor4
                        !Desvalor5 = ZDesValor5
                        !Desvalor6 = ZDesValor6
                        !Desvalor7 = ZDesValor7
                        !Desvalor8 = ZDesValor8
                        !Desvalor9 = ZDesValor9
                        !Desvalor10 = ZDesValor10
                        !Desvalor11 = ZDesValor11
                        !Desvalor12 = ZDesValor12
                        !Desvalor13 = ZDesValor13
                        !Desvalor14 = ZDesValor14
                        !Desvalor15 = ZDesValor15
                        !Desvalor16 = ZDesValor16
                        !Desvalor17 = ZDesValor17
                        !Desvalor18 = ZDesValor18
                        !DesValor19 = ZDesValor19
                        !Desvalor20 = ZDesValor20
                        !Desvalor21 = ZDesValor21
                        !Desvalor22 = ZDesValor22
                        !Desvalor23 = ZDesValor23
                        !Desvalor24 = ZDesValor24
                        !Desvalor25 = ZDesValor25
                        !Desvalor26 = ZDesValor26
                        !Desvalor27 = ZDesValor27
                        !Desvalor28 = ZDesValor28
                        !Desvalor29 = ZDesValor29
                        !Desvalor30 = ZDesValor30
                        !DesProducto = ZDesProducto
                        .Update
                    End With
                
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
            
        End With
        
        rstPrueart.Close
        
    End If

    lista.ReportFileName = "WPruArt.rpt"
    
    lista.WindowTitle = "Listado de Controles de Materias Primas"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height
    
    Rem Uno = "{Prueart.Fechaord} in " + Chr$(34) + FDesde + Chr$(34) + " to " + Chr$(34) + FHasta + Chr$(34)
    Rem Dos = " and {Prueart.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Desde.Text + Chr$(34)
    Rem lista.GroupSelectionFormula = Uno + Dos
    
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    
    lista.DataFiles(0) = WEmpresa + "auxi.mdb"
    lista.DataFiles(1) = WEmpresa + "auxi.mdb"
    lista.DataFiles(2) = ""
    lista.DataFiles(3) = ""
    
    lista.Connect = Connect()
    
    lista.Action = 1
    Frame2.Visible = False
End Sub


Private Sub btnConsultaPartida_Click()
    PrgConsultaPartidaDy.Show
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub CancelaLote_Click()
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub cmdAddlote_Click()
    Rem With rstprueart
    Rem     .Index = "Prueba"
    Rem     ClavePrue$ = "1999999"
    Rem     .Seek "<", ClavePrue$
    Rem     If .NoMatch Then
    Rem         Lote.Text = "1"
    Rem             Else
    Rem         Lote.Text = Str$(Val(!Prueba) + 1)
    Rem     End If
    Rem
    Rem     Auxi1 = Lote.Text
    Rem     Call Ceros(Auxi1, 6)
    Rem     Lote.Text = Auxi1
    Rem
    Rem     Auxi = "1"
    Rem
    Rem     Liberada.Text = ""
    Rem     Devuelta.Text = ""
    Rem     NroRechazo.Text = ""
    Rem     Nueva.Text = ""
    Rem
    Rem     panLote.Visible = True
    Rem
    Rem     Liberada.SetFocus
    Rem
    Rem End With
    
    Auxi = "1"
    
    Lote.Text = ""
    Liberada.Text = ""
    Devuelta.Text = ""
    NroRechazo.Text = ""
    Nueva.Text = ""
    PartidaProveedor.Text = ""
    OrigenMercaderia.Text = RTrim(WOrigen)
    
    Select Case Val(WEmpresa)
        Case 3
            ZSql = ""
            ZSql = ZSql + "Select Clave,Laudo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo < " + "'" + "230000" + "'"
            ZSql = ZSql + " Order by Clave"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.MoveLast
                Lote.Text = Trim(Str$(rstLaudo!Laudo + 1))
                rstLaudo.Close
            End If
            
        Case 5
            ZSql = ""
            ZSql = ZSql + "Select Clave,Laudo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo < " + "'" + "389999" + "'"
            ZSql = ZSql + " Order by Clave"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.MoveLast
                Lote.Text = Trim(Str$(rstLaudo!Laudo + 1))
                rstLaudo.Close
            End If
            
        Case 6
            ZSql = ""
            ZSql = ZSql + "Select Clave,Laudo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo < " + "'" + "489999" + "'"
            ZSql = ZSql + " Order by Clave"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.MoveLast
                Lote.Text = Trim(Str$(rstLaudo!Laudo + 1))
                rstLaudo.Close
            End If
            
        Case 7
            ZSql = ""
            ZSql = ZSql + "Select Clave,Laudo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo < " + "'" + "519999" + "'"
            ZSql = ZSql + " Order by Clave"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.MoveLast
                Lote.Text = Trim(Str$(rstLaudo!Laudo + 1))
                rstLaudo.Close
            End If
            
        Case 10
            ZSql = ""
            ZSql = ZSql + "Select Clave,Laudo"
            ZSql = ZSql + " FROM Laudo"
            ZSql = ZSql + " Where Laudo < " + "'" + "539999" + "'"
            ZSql = ZSql + " Order by Clave"
            spLaudo = ZSql
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                rstLaudo.MoveLast
                Lote.Text = Trim(Str$(rstLaudo!Laudo + 1))
                rstLaudo.Close
            End If
        
        Case Else
    
    End Select
        
    panLote.Visible = True
    
    Lote.SetFocus
        
End Sub

Private Sub cmdAddRechazo_Click()

    spPrueart = "ConsultaPruebaMenor " + "'" + "2999999" + "'"
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        If rstPrueart!Tipo = "2" Then
            Lote.Text = Str$(Val(rstPrueart!Prueba) + 1)
                Else
            Lote.Text = "2000"
        End If
        rstPrueart.Close
            Else
        Lote.Text = "1"
    End If
    
    Auxi1 = Lote.Text
    Call Ceros(Auxi1, 6)
    Lote.Text = Auxi1
        
    Auxi = "2"
            
    Liberada.Text = ""
    Devuelta.Text = ""
    NroRechazo.Text = ""
    Nueva.Text = ""
        
    panLote.Visible = True
        
    Liberada.SetFocus
        
End Sub


Private Sub Desvio_Click()
    If Val(Partida.Text) <> 0 Then
        ZLoteRevalida = Partida.Text
        ZFechaRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZArticuloRevalida = Producto.Text
        ZDesArticuloRevalida = Descriprod.Caption
        ZFechaVencimiento = Vto.Text
        PrgDesvio.Show
    End If
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
    OPEN_FILE_PRUEBA
End Sub

Private Sub GrabaLote_Click()



    On Error GoTo WError
    
    Entra = "N"

    If Val(Liberada.Text) > 0 Then

        Select Case Val(WEmpresa)
            Case 1
                If Val(Lote.Text) >= 100000 And Val(Lote.Text) <= 189999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 190000 And Val(Lote.Text) <= 194999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 900000 And Val(Lote.Text) <= 989999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 990000 And Val(Lote.Text) <= 994999 Then
                    Entra = "S"
                End If
            Case 2
                If Val(Lote.Text) >= 600000 And Val(Lote.Text) <= 649999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 690000 And Val(Lote.Text) <= 691999 Then
                    Entra = "S"
                End If
            Case 3
                If Val(Lote.Text) >= 200000 And Val(Lote.Text) <= 289999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 290000 And Val(Lote.Text) <= 294999 Then
                    Entra = "S"
                End If
            Case 4
                If Val(Lote.Text) >= 700000 And Val(Lote.Text) <= 789999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 790000 And Val(Lote.Text) <= 794999 Then
                    Entra = "S"
                End If
            Case 5
                If Val(Lote.Text) >= 300000 And Val(Lote.Text) <= 389999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 390000 And Val(Lote.Text) <= 394999 Then
                    Entra = "S"
                End If
            Case 6
                If Val(Lote.Text) >= 400000 And Val(Lote.Text) <= 489999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 490000 And Val(Lote.Text) <= 494999 Then
                    Entra = "S"
                End If
            Case 7
                If Val(Lote.Text) >= 500000 And Val(Lote.Text) <= 519999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 590000 And Val(Lote.Text) <= 594999 Then
                    Entra = "S"
                End If
            Case 8
                If Val(Lote.Text) >= 800000 And Val(Lote.Text) <= 889999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 890000 And Val(Lote.Text) <= 894999 Then
                    Entra = "S"
                End If
            Case 9
                If Val(Lote.Text) >= 650000 And Val(Lote.Text) <= 689999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 692000 And Val(Lote.Text) <= 694999 Then
                    Entra = "S"
                End If
            Case 10
                If Val(Lote.Text) >= 520000 And Val(Lote.Text) <= 539999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 570000 And Val(Lote.Text) <= 574999 Then
                    Entra = "S"
                End If
            Case 11
                If Val(Lote.Text) >= 540000 And Val(Lote.Text) <= 559999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 580000 And Val(Lote.Text) <= 584999 Then
                    Entra = "S"
                End If
            Case Else
                Entra = "S"
        End Select
    End If
        
    If Val(Devuelta.Text) > 0 Then
    
        Select Case Val(WEmpresa)
            Case 1
                If Val(Lote.Text) >= 70000 And Val(Lote.Text) <= 70999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 995000 And Val(Lote.Text) <= 999999 Then
                    Entra = "S"
                End If
            Case 2
                If Val(Lote.Text) >= 71000 And Val(Lote.Text) <= 71299 Then
                    Entra = "S"
                End If
            Case 3
                If Val(Lote.Text) >= 74000 And Val(Lote.Text) <= 74999 Then
                    Entra = "S"
                End If
            Case 4
                If Val(Lote.Text) >= 76000 And Val(Lote.Text) <= 76999 Then
                    Entra = "S"
                End If
            Case 5
                If Val(Lote.Text) >= 78000 And Val(Lote.Text) <= 78999 Then
                    Entra = "S"
                End If
            Case 6
                If Val(Lote.Text) >= 75000 And Val(Lote.Text) <= 75999 Then
                    Entra = "S"
                End If
            Case 7
                If Val(Lote.Text) >= 72000 And Val(Lote.Text) <= 72399 Then
                    Entra = "S"
                End If
            Case 8
                If Val(Lote.Text) >= 73000 And Val(Lote.Text) <= 73999 Then
                    Entra = "S"
                End If
            Case 9
                If Val(Lote.Text) >= 71300 And Val(Lote.Text) <= 71999 Then
                    Entra = "S"
                End If
            Case 10
                If Val(Lote.Text) >= 72400 And Val(Lote.Text) <= 72699 Then
                    Entra = "S"
                End If
            Case 11
                If Val(Lote.Text) >= 72700 And Val(Lote.Text) <= 72999 Then
                    Entra = "S"
                End If
            Case Else
                Entra = "S"
        End Select
        
    End If
    
    If Entra = "N" Then
        PrgPrueArtRango.Show
        Exit Sub
    End If
    
    spLaudo = "ListaLaudo " + "'" + Lote.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        m$ = "Numero de lote ya existente"
        A% = MsgBox(m$, 0, "Pruebas de Materias Primas")
        rstLaudo.Close
        Exit Sub
    End If

    If Val(Liberada.Text) = 0 Then
        Liberada.Text = "0"
    End If
    
    If Val(Devuelta.Text) = 0 Then
        Devuelta.Text = "0"
    End If
    
    If Val(NroRechazo.Text) = 0 Then
        NroRechazo.Text = "0"
    End If
    
    Cantidad = Val(Liberada.Text) + Val(Devuelta.Text)

    Call Calcula_SaldoOrden
    
    Rem dada
    Rem ojo
    Rem sacar esto
    Rem SaldoOrden = 1000

    If Cantidad > SaldoOrden Then
        m$ = "La Cantidad supera al saldo del informe de recepcion" + Chr$(13) _
             + "Cantidad recibida (Informe de recepcion) : " + Str$(WRecibida) + Chr$(13) _
             + "Cantidad Laudada (Laudos Anteriores) : " + Str$(WLaudada) + Chr$(13) _
             + "Saldo Disponible para laudar : " + Str$(SaldoOrden)
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Exit Sub
            Else
        If Cantidad < SaldoOrden Then
            m$ = "Atencion : la cantidad laudada es menor al saldo del informe de recepcion" + Chr$(13) _
                + "Cantidad recibida (Informe de recepcion) : " + Str$(WRecibida) + Chr$(13) _
                + "Cantidad Laudada (Laudos Anteriores) : " + Str$(WLaudada) + Chr$(13) _
                + "Saldo Disponible para laudar : " + Str$(SaldoOrden) + Chr$(13) _
                + "Saldo Pendiente para futuros laudos : " + Str$(SaldoOrden - Cantidad) + Chr$(13) _
                + "Confirma la grabacion del LAUDO"
                Respuesta% = MsgBox(m$, 32 + 4, "Ingreso de Pruebas")
                If Respuesta% = 7 Then
                    Exit Sub
                End If
        End If
    End If
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Erase ZDesde
    Erase ZHasta

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        ZDesde(1) = IIf(IsNull(rstEspecificacionesUnifica!Desde1), "", rstEspecificacionesUnifica!Desde1)
        ZDesde(2) = IIf(IsNull(rstEspecificacionesUnifica!Desde2), "", rstEspecificacionesUnifica!Desde2)
        ZDesde(3) = IIf(IsNull(rstEspecificacionesUnifica!Desde3), "", rstEspecificacionesUnifica!Desde3)
        ZDesde(4) = IIf(IsNull(rstEspecificacionesUnifica!Desde4), "", rstEspecificacionesUnifica!Desde4)
        ZDesde(5) = IIf(IsNull(rstEspecificacionesUnifica!Desde5), "", rstEspecificacionesUnifica!Desde5)
        ZDesde(6) = IIf(IsNull(rstEspecificacionesUnifica!Desde6), "", rstEspecificacionesUnifica!Desde6)
        ZDesde(7) = IIf(IsNull(rstEspecificacionesUnifica!Desde7), "", rstEspecificacionesUnifica!Desde7)
        ZDesde(8) = IIf(IsNull(rstEspecificacionesUnifica!Desde8), "", rstEspecificacionesUnifica!Desde8)
        ZDesde(9) = IIf(IsNull(rstEspecificacionesUnifica!Desde9), "", rstEspecificacionesUnifica!Desde9)
        ZDesde(10) = IIf(IsNull(rstEspecificacionesUnifica!Desde10), "", rstEspecificacionesUnifica!Desde10)
        ZDesde(11) = IIf(IsNull(rstEspecificacionesUnifica!Desde11), "", rstEspecificacionesUnifica!Desde11)
        ZDesde(12) = IIf(IsNull(rstEspecificacionesUnifica!Desde12), "", rstEspecificacionesUnifica!Desde12)
        ZDesde(13) = IIf(IsNull(rstEspecificacionesUnifica!Desde13), "", rstEspecificacionesUnifica!Desde13)
        ZDesde(14) = IIf(IsNull(rstEspecificacionesUnifica!Desde14), "", rstEspecificacionesUnifica!Desde14)
        ZDesde(15) = IIf(IsNull(rstEspecificacionesUnifica!Desde15), "", rstEspecificacionesUnifica!Desde15)
        ZDesde(16) = IIf(IsNull(rstEspecificacionesUnifica!Desde16), "", rstEspecificacionesUnifica!Desde16)
        ZDesde(17) = IIf(IsNull(rstEspecificacionesUnifica!Desde17), "", rstEspecificacionesUnifica!Desde17)
        ZDesde(18) = IIf(IsNull(rstEspecificacionesUnifica!Desde18), "", rstEspecificacionesUnifica!Desde18)
        ZDesde(19) = IIf(IsNull(rstEspecificacionesUnifica!Desde19), "", rstEspecificacionesUnifica!Desde19)
        ZDesde(20) = IIf(IsNull(rstEspecificacionesUnifica!Desde20), "", rstEspecificacionesUnifica!Desde20)
        
        ZHasta(1) = IIf(IsNull(rstEspecificacionesUnifica!Hasta1), "", rstEspecificacionesUnifica!Hasta1)
        ZHasta(2) = IIf(IsNull(rstEspecificacionesUnifica!Hasta2), "", rstEspecificacionesUnifica!Hasta2)
        ZHasta(3) = IIf(IsNull(rstEspecificacionesUnifica!Hasta3), "", rstEspecificacionesUnifica!Hasta3)
        ZHasta(4) = IIf(IsNull(rstEspecificacionesUnifica!Hasta4), "", rstEspecificacionesUnifica!Hasta4)
        ZHasta(5) = IIf(IsNull(rstEspecificacionesUnifica!Hasta5), "", rstEspecificacionesUnifica!Hasta5)
        ZHasta(6) = IIf(IsNull(rstEspecificacionesUnifica!Hasta6), "", rstEspecificacionesUnifica!Hasta6)
        ZHasta(7) = IIf(IsNull(rstEspecificacionesUnifica!Hasta7), "", rstEspecificacionesUnifica!Hasta7)
        ZHasta(8) = IIf(IsNull(rstEspecificacionesUnifica!Hasta8), "", rstEspecificacionesUnifica!Hasta8)
        ZHasta(9) = IIf(IsNull(rstEspecificacionesUnifica!Hasta9), "", rstEspecificacionesUnifica!Hasta9)
        ZHasta(10) = IIf(IsNull(rstEspecificacionesUnifica!Hasta10), "", rstEspecificacionesUnifica!Hasta10)
        ZHasta(11) = IIf(IsNull(rstEspecificacionesUnifica!Hasta11), "", rstEspecificacionesUnifica!Hasta11)
        ZHasta(12) = IIf(IsNull(rstEspecificacionesUnifica!Hasta12), "", rstEspecificacionesUnifica!Hasta12)
        ZHasta(13) = IIf(IsNull(rstEspecificacionesUnifica!Hasta13), "", rstEspecificacionesUnifica!Hasta13)
        ZHasta(14) = IIf(IsNull(rstEspecificacionesUnifica!Hasta14), "", rstEspecificacionesUnifica!Hasta14)
        ZHasta(15) = IIf(IsNull(rstEspecificacionesUnifica!Hasta15), "", rstEspecificacionesUnifica!Hasta15)
        ZHasta(16) = IIf(IsNull(rstEspecificacionesUnifica!Hasta16), "", rstEspecificacionesUnifica!Hasta16)
        ZHasta(17) = IIf(IsNull(rstEspecificacionesUnifica!Hasta17), "", rstEspecificacionesUnifica!Hasta17)
        ZHasta(18) = IIf(IsNull(rstEspecificacionesUnifica!Hasta18), "", rstEspecificacionesUnifica!Hasta18)
        ZHasta(19) = IIf(IsNull(rstEspecificacionesUnifica!Hasta19), "", rstEspecificacionesUnifica!Hasta19)
        ZHasta(20) = IIf(IsNull(rstEspecificacionesUnifica!Hasta20), "", rstEspecificacionesUnifica!Hasta20)
        
        ZDesde(1) = Trim(ZDesde(1))
        ZDesde(2) = Trim(ZDesde(2))
        ZDesde(3) = Trim(ZDesde(3))
        ZDesde(4) = Trim(ZDesde(4))
        ZDesde(5) = Trim(ZDesde(5))
        ZDesde(6) = Trim(ZDesde(6))
        ZDesde(7) = Trim(ZDesde(7))
        ZDesde(8) = Trim(ZDesde(8))
        ZDesde(9) = Trim(ZDesde(9))
        ZDesde(10) = Trim(ZDesde(10))
        ZDesde(11) = Trim(ZDesde(11))
        ZDesde(12) = Trim(ZDesde(12))
        ZDesde(13) = Trim(ZDesde(13))
        ZDesde(14) = Trim(ZDesde(14))
        ZDesde(15) = Trim(ZDesde(15))
        ZDesde(16) = Trim(ZDesde(16))
        ZDesde(17) = Trim(ZDesde(17))
        ZDesde(18) = Trim(ZDesde(18))
        ZDesde(19) = Trim(ZDesde(19))
        ZDesde(20) = Trim(ZDesde(20))
        
        ZHasta(1) = Trim(ZHasta(1))
        ZHasta(2) = Trim(ZHasta(2))
        ZHasta(3) = Trim(ZHasta(3))
        ZHasta(4) = Trim(ZHasta(4))
        ZHasta(5) = Trim(ZHasta(5))
        ZHasta(6) = Trim(ZHasta(6))
        ZHasta(7) = Trim(ZHasta(7))
        ZHasta(8) = Trim(ZHasta(8))
        ZHasta(9) = Trim(ZHasta(9))
        ZHasta(10) = Trim(ZHasta(10))
        ZHasta(11) = Trim(ZHasta(11))
        ZHasta(12) = Trim(ZHasta(12))
        ZHasta(13) = Trim(ZHasta(13))
        ZHasta(14) = Trim(ZHasta(14))
        ZHasta(15) = Trim(ZHasta(15))
        ZHasta(16) = Trim(ZHasta(16))
        ZHasta(17) = Trim(ZHasta(17))
        ZHasta(18) = Trim(ZHasta(18))
        ZHasta(19) = Trim(ZHasta(19))
        ZHasta(20) = Trim(ZHasta(20))
        
        ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        rstEspecificacionesUnifica.Close
    End If
    

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Producto.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        ZDesde(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde21), "", rstEspecificacionesUnificaIII!Desde21)
        ZDesde(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde22), "", rstEspecificacionesUnificaIII!Desde22)
        ZDesde(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde23), "", rstEspecificacionesUnificaIII!Desde23)
        ZDesde(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde24), "", rstEspecificacionesUnificaIII!Desde24)
        ZDesde(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde25), "", rstEspecificacionesUnificaIII!Desde25)
        ZDesde(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde26), "", rstEspecificacionesUnificaIII!Desde26)
        ZDesde(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde27), "", rstEspecificacionesUnificaIII!Desde27)
        ZDesde(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde28), "", rstEspecificacionesUnificaIII!Desde28)
        ZDesde(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde29), "", rstEspecificacionesUnificaIII!Desde29)
        ZDesde(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde30), "", rstEspecificacionesUnificaIII!Desde30)
        
        ZHasta(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta21), "", rstEspecificacionesUnificaIII!Hasta21)
        ZHasta(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta22), "", rstEspecificacionesUnificaIII!Hasta22)
        ZHasta(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta23), "", rstEspecificacionesUnificaIII!Hasta23)
        ZHasta(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta24), "", rstEspecificacionesUnificaIII!Hasta24)
        ZHasta(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta25), "", rstEspecificacionesUnificaIII!Hasta25)
        ZHasta(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta26), "", rstEspecificacionesUnificaIII!Hasta26)
        ZHasta(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta27), "", rstEspecificacionesUnificaIII!Hasta27)
        ZHasta(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta28), "", rstEspecificacionesUnificaIII!Hasta28)
        ZHasta(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta29), "", rstEspecificacionesUnificaIII!Hasta29)
        ZHasta(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta30), "", rstEspecificacionesUnificaIII!Hasta30)
        
        ZDesde(21) = Trim(ZDesde(21))
        ZDesde(22) = Trim(ZDesde(22))
        ZDesde(23) = Trim(ZDesde(23))
        ZDesde(24) = Trim(ZDesde(24))
        ZDesde(25) = Trim(ZDesde(25))
        ZDesde(26) = Trim(ZDesde(26))
        ZDesde(27) = Trim(ZDesde(27))
        ZDesde(28) = Trim(ZDesde(28))
        ZDesde(29) = Trim(ZDesde(29))
        ZDesde(30) = Trim(ZDesde(30))
        
        ZHasta(21) = Trim(ZHasta(21))
        ZHasta(22) = Trim(ZHasta(22))
        ZHasta(23) = Trim(ZHasta(23))
        ZHasta(24) = Trim(ZHasta(24))
        ZHasta(25) = Trim(ZHasta(25))
        ZHasta(26) = Trim(ZHasta(26))
        ZHasta(27) = Trim(ZHasta(27))
        ZHasta(28) = Trim(ZHasta(28))
        ZHasta(29) = Trim(ZHasta(29))
        ZHasta(30) = Trim(ZHasta(30))
        
        ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        rstEspecificacionesUnificaIII.Close
    End If
    
    
    Call Conecta_Empresa
    
    ZValorNumero(1) = ValorNumero1.Text
    ZValorNumero(2) = ValorNumero2.Text
    ZValorNumero(3) = ValorNumero3.Text
    ZValorNumero(4) = ValorNumero4.Text
    ZValorNumero(5) = ValorNumero5.Text
    ZValorNumero(6) = ValorNumero6.Text
    ZValorNumero(7) = ValorNumero7.Text
    ZValorNumero(8) = ValorNumero8.Text
    ZValorNumero(9) = ValorNumero9.Text
    ZValorNumero(10) = ValorNumero10.Text
    ZValorNumero(11) = ValorNumero11.Text
    ZValorNumero(12) = ValorNumero12.Text
    ZValorNumero(13) = ValorNumero13.Text
    ZValorNumero(14) = ValorNumero14.Text
    ZValorNumero(15) = ValorNumero15.Text
    ZValorNumero(16) = ValorNumero16.Text
    ZValorNumero(17) = ValorNumero17.Text
    ZValorNumero(18) = ValorNumero18.Text
    ZValorNumero(19) = ValorNumero19.Text
    ZValorNumero(20) = ValorNumero20.Text
    ZValorNumero(21) = ValorNumero21.Text
    ZValorNumero(22) = ValorNumero22.Text
    ZValorNumero(23) = ValorNumero23.Text
    ZValorNumero(24) = ValorNumero24.Text
    ZValorNumero(25) = ValorNumero25.Text
    ZValorNumero(26) = ValorNumero26.Text
    ZValorNumero(27) = ValorNumero27.Text
    ZValorNumero(28) = ValorNumero28.Text
    ZValorNumero(29) = ValorNumero29.Text
    ZValorNumero(30) = ValorNumero30.Text
    
    Rem For WWCiclo = 1 To 20
    Rem     If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
    Rem         If Val(ZValorNumero(WWCiclo)) = 0 Then
    Rem             m$ = "No se informado valor de control en una de las ensayos que requiere validacion"
    Rem             A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    Rem             Exit Sub
    Rem         End If
    Rem     End If
    Rem Next WWCiclo
    
    ZDesvio = "S"
    
    Select Case Val(WEmpresa)
        Case 1
            If Val(Lote.Text) >= 100000 And Val(Lote.Text) <= 189999 Then
                ZDesvio = "N"
            End If
            If Val(Lote.Text) >= 900000 And Val(Lote.Text) <= 989999 Then
                ZDesvio = "N"
            End If
        Case 3
            If Val(Lote.Text) >= 200000 And Val(Lote.Text) <= 289999 Then
                ZDesvio = "N"
            End If
        Case 5
            If Val(Lote.Text) >= 300000 And Val(Lote.Text) <= 389999 Then
                ZDesvio = "N"
            End If
        Case 6
            If Val(Lote.Text) >= 400000 And Val(Lote.Text) <= 489999 Then
                ZDesvio = "N"
            End If
        Case 7, 10, 11
            If Val(Lote.Text) >= 500000 And Val(Lote.Text) <= 559999 Then
                ZDesvio = "N"
            End If
        Case 2
            If Val(Lote.Text) >= 600000 And Val(Lote.Text) <= 689999 Then
                ZDesvio = "N"
            End If
        Case 4
            If Val(Lote.Text) >= 700000 And Val(Lote.Text) <= 789999 Then
                ZDesvio = "N"
            End If
        Case 8
            If Val(Lote.Text) >= 800000 And Val(Lote.Text) <= 889999 Then
                ZDesvio = "N"
            End If
    End Select
    
    ZCategoriaI = 0
    WProveedor = ""
    
    Sql1 = "Select * "
    Sql2 = " FROM Orden"
    Sql3 = " Where Orden = " + "'" + Orden.Text + "'"
    spOrden = Sql1 + Sql2 + Sql3
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        WProveedor = rstOrden!Proveedor
        rstOrden.Close
    End If
    
    
    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        ZCategoriaI = IIf(IsNull(RstProveedor!CategoriaI), "0", RstProveedor!CategoriaI)
        RstProveedor.Close
    End If
    
    Call Conecta_Empresa
    
    If Left$(Producto.Text, 2) = "DQ" Then
        ZCategoriaI = 2
    End If
    
    If ZDesvio = "N" Then
    
        If ZCategoriaI <> 1 Then
    
            For WWCiclo = 1 To 30
            
                If Val(ZEnsayo(WWCiclo)) <> 0 Then
            
                    If Val(ZDesde(WWCiclo)) <> 0 Or Val(ZHasta(WWCiclo)) <> 0 Then
                    
                        If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                            aa = Val(ZValorNumero(WWCiclo))
                            If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Or Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                                    
                        If Val(ZDesde(WWCiclo)) <> 0 And Val(ZHasta(WWCiclo)) = 0 Then
                            If Val(ZValorNumero(WWCiclo)) < Val(ZDesde(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                                
                        If Val(ZDesde(WWCiclo)) = 0 And Val(ZHasta(WWCiclo)) <> 0 Then
                            If Val(ZValorNumero(WWCiclo)) > Val(ZHasta(WWCiclo)) Then
                                m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                                Exit Sub
                            End If
                        End If
                        
                            Else
                            
                        If Trim(UCase(ZValorNumero(WWCiclo))) <> "S" Then
                            m$ = "El valor de uno de los resultados de las pruebas realizadas no concuerda con los valores permitidos"
                            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                            Exit Sub
                        End If
                        
                    End If
                    
                End If
            
            Next WWCiclo
                
                Else
                    
            Ensayo.Text = "Proveedor A - Aprobado con Cert.Analisis"
            
            Rem For WWCiclo = 1 To 20
            Rem
            Rem     If ZEnsayo(WWCiclo) = 40 Then
            Rem
            Rem         select cASe
            Rem
            Rem     End If
            Rem
            Rem Next WWCiclo
            
        End If
        
    End If
    
    XEnvase = 0
    
    ZSql = ""
    ZSql = ZSql + "Select * "
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
    ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        Informe = rstInforme!Informe
        XEnvase = rstInforme!Envase
        XCertificado1 = IIf(IsNull(rstInforme!Certificado1), "0", rstInforme!Certificado1)
        XCertificado2 = IIf(IsNull(rstInforme!Certificado2), "", RTrim(rstInforme!Certificado2))
        XEstado1 = IIf(IsNull(rstInforme!Estado1), "0", rstInforme!Estado1)
        XEstado2 = IIf(IsNull(rstInforme!Estado2), "", RTrim(rstInforme!Estado2))
        ZVencimiento = IIf(IsNull(rstInforme!fechavencimiento), "  /  /    ", rstInforme!fechavencimiento)
        rstInforme.Close
    End If
    
    WInforme = Informe
    
    If Val(Liberada.Text) > 0 Then
    
        Auxi = "1"
    
        WLote = Lote.Text
        Call Ceros(WLote, 6)
        WPrueba = Auxi + WLote
        WProducto = Producto.Text
        WFecha = fecha.Text
        WOrden = Orden.Text
        
        WValor1 = Valor1.Text
        WValor2 = Valor2.Text
        WValor3 = Valor3.Text
        WValor4 = Valor4.Text
        WValor5 = Valor5.Text
        WValor6 = Valor6.Text
        WValor7 = Valor7.Text
        WValor8 = Valor8.Text
        WValor9 = Valor9.Text
        WValor10 = Valor10.Text
        WValor11 = Valor11.Text
        WValor12 = Valor12.Text
        WValor13 = Valor13.Text
        WValor14 = Valor14.Text
        WValor15 = Valor15.Text
        WValor16 = Valor16.Text
        WValor17 = Valor17.Text
        WValor18 = Valor18.Text
        WValor19 = Valor19.Text
        WValor20 = Valor20.Text
        WValor21 = Valor21.Text
        WValor22 = Valor22.Text
        WValor23 = Valor23.Text
        WValor24 = Valor24.Text
        WValor25 = Valor25.Text
        WValor26 = Valor26.Text
        WValor27 = Valor27.Text
        WValor28 = Valor28.Text
        WValor29 = Valor29.Text
        WValor30 = Valor30.Text
        
        WValorNumero1 = ValorNumero1.Text
        WValorNumero2 = ValorNumero2.Text
        WValorNumero3 = ValorNumero3.Text
        WValorNumero4 = ValorNumero4.Text
        WValorNumero5 = ValorNumero5.Text
        WValorNumero6 = ValorNumero6.Text
        WValorNumero7 = ValorNumero7.Text
        WValorNumero8 = ValorNumero8.Text
        WValorNumero9 = ValorNumero9.Text
        WValorNumero10 = ValorNumero10.Text
        WValorNumero11 = ValorNumero11.Text
        WValorNumero12 = ValorNumero12.Text
        WValorNumero13 = ValorNumero13.Text
        WValorNumero14 = ValorNumero14.Text
        WValorNumero15 = ValorNumero15.Text
        WValorNumero16 = ValorNumero16.Text
        WValorNumero17 = ValorNumero17.Text
        WValorNumero18 = ValorNumero18.Text
        WValorNumero19 = ValorNumero19.Text
        WValorNumero20 = ValorNumero20.Text
        WValorNumero21 = ValorNumero21.Text
        WValorNumero22 = ValorNumero22.Text
        WValorNumero23 = ValorNumero23.Text
        WValorNumero24 = ValorNumero24.Text
        WValorNumero25 = ValorNumero25.Text
        WValorNumero26 = ValorNumero26.Text
        WValorNumero27 = ValorNumero27.Text
        WValorNumero28 = ValorNumero28.Text
        WValorNumero29 = ValorNumero29.Text
        WValorNumero30 = ValorNumero30.Text
        
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = Liberada.Text
        WDevuelta = "0"
        WLote = Lote.Text
        WRechazo = ""
        WNueva = "N"
        WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
        WDate = Date$
        WObserva2 = ""
    
        Rem XParam = "'" + WPrueba + "','" _
        rem         + WProducto + "','" _
        rem         + WFecha + "','" _
        rem         + WOrden + "','" _
        rem         + WValor1 + "','" _
        rem         + WValor2 + "','" _
        rem         + WValor3 + "','" _
        rem         + WValor4 + "','" _
        rem         + WValor5 + "','" _
        rem         + WValor6 + "','" _
        rem         + WValor7 + "','" _
        rem         + WValor8 + "','" _
        rem         + WValor9 + "','" _
        rem         + WValor10 + "','" _
        rem         + WEnsayo + "','" _
        rem         + WAspecto + "','" _
        rem         + WObservaciones + "','" _
        rem         + WObserva2 + "','" _
        rem         + WConfecciono + "','" _
        rem         + WLiberada + "','" _
        rem         + WDevuelta + "','" _
        rem         + WLote + "','" _
        rem         + WRechazo + "','" _
        rem         + WNueva + "','" + WFechaOrd + "','" _
        rem         + WDate + "'"
        Rem Set rstPrueart = db.OpenRecordset("AltaPrueart " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO PrueArt ("
        ZSql = ZSql + "Prueba ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "Valor11 ,"
        ZSql = ZSql + "Valor12 ,"
        ZSql = ZSql + "Valor13 ,"
        ZSql = ZSql + "Valor14 ,"
        ZSql = ZSql + "Valor15 ,"
        ZSql = ZSql + "Valor16 ,"
        ZSql = ZSql + "Valor17 ,"
        ZSql = ZSql + "Valor18 ,"
        ZSql = ZSql + "Valor19 ,"
        ZSql = ZSql + "Valor20 ,"
        ZSql = ZSql + "Valor21 ,"
        ZSql = ZSql + "Valor22 ,"
        ZSql = ZSql + "Valor23 ,"
        ZSql = ZSql + "Valor24 ,"
        ZSql = ZSql + "Valor25 ,"
        ZSql = ZSql + "Valor26 ,"
        ZSql = ZSql + "Valor27 ,"
        ZSql = ZSql + "Valor28 ,"
        ZSql = ZSql + "Valor29 ,"
        ZSql = ZSql + "Valor30 ,"
        ZSql = ZSql + "ValorNumero1 ,"
        ZSql = ZSql + "ValorNumero2 ,"
        ZSql = ZSql + "ValorNumero3 ,"
        ZSql = ZSql + "ValorNumero4 ,"
        ZSql = ZSql + "ValorNumero5 ,"
        ZSql = ZSql + "ValorNumero6 ,"
        ZSql = ZSql + "ValorNumero7 ,"
        ZSql = ZSql + "ValorNumero8 ,"
        ZSql = ZSql + "ValorNumero9 ,"
        ZSql = ZSql + "ValorNumero10 ,"
        ZSql = ZSql + "ValorNumero11 ,"
        ZSql = ZSql + "ValorNumero12 ,"
        ZSql = ZSql + "ValorNumero13 ,"
        ZSql = ZSql + "ValorNumero14 ,"
        ZSql = ZSql + "ValorNumero15 ,"
        ZSql = ZSql + "ValorNumero16 ,"
        ZSql = ZSql + "ValorNumero17 ,"
        ZSql = ZSql + "ValorNumero18 ,"
        ZSql = ZSql + "ValorNumero19 ,"
        ZSql = ZSql + "ValorNumero20 ,"
        ZSql = ZSql + "ValorNumero21 ,"
        ZSql = ZSql + "ValorNumero22 ,"
        ZSql = ZSql + "ValorNumero23 ,"
        ZSql = ZSql + "ValorNumero24 ,"
        ZSql = ZSql + "ValorNumero25 ,"
        ZSql = ZSql + "ValorNumero26 ,"
        ZSql = ZSql + "ValorNumero27 ,"
        ZSql = ZSql + "ValorNumero28 ,"
        ZSql = ZSql + "ValorNumero29 ,"
        ZSql = ZSql + "ValorNumero30 ,"
        ZSql = ZSql + "Ensayo ,"
        ZSql = ZSql + "Aspecto ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Observa2 ,"
        ZSql = ZSql + "Confecciono ,"
        ZSql = ZSql + "Liberada ,"
        ZSql = ZSql + "Devuelta ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Rechazo ,"
        ZSql = ZSql + "Nueva ,"
        ZSql = ZSql + "FechaOrd ,"
        ZSql = ZSql + "WDate )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WPrueba + "',"
        ZSql = ZSql + "'" + WProducto + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WOrden + "',"
        ZSql = ZSql + "'" + WValor1 + "',"
        ZSql = ZSql + "'" + WValor2 + "',"
        ZSql = ZSql + "'" + WValor3 + "',"
        ZSql = ZSql + "'" + WValor4 + "',"
        ZSql = ZSql + "'" + WValor5 + "',"
        ZSql = ZSql + "'" + WValor6 + "',"
        ZSql = ZSql + "'" + WValor7 + "',"
        ZSql = ZSql + "'" + WValor8 + "',"
        ZSql = ZSql + "'" + WValor9 + "',"
        ZSql = ZSql + "'" + WValor10 + "',"
        ZSql = ZSql + "'" + WValor11 + "',"
        ZSql = ZSql + "'" + WValor12 + "',"
        ZSql = ZSql + "'" + WValor13 + "',"
        ZSql = ZSql + "'" + WValor14 + "',"
        ZSql = ZSql + "'" + WValor15 + "',"
        ZSql = ZSql + "'" + WValor16 + "',"
        ZSql = ZSql + "'" + WValor17 + "',"
        ZSql = ZSql + "'" + WValor18 + "',"
        ZSql = ZSql + "'" + WValor19 + "',"
        ZSql = ZSql + "'" + WValor20 + "',"
        ZSql = ZSql + "'" + WValor21 + "',"
        ZSql = ZSql + "'" + WValor22 + "',"
        ZSql = ZSql + "'" + WValor23 + "',"
        ZSql = ZSql + "'" + WValor24 + "',"
        ZSql = ZSql + "'" + WValor25 + "',"
        ZSql = ZSql + "'" + WValor26 + "',"
        ZSql = ZSql + "'" + WValor27 + "',"
        ZSql = ZSql + "'" + WValor28 + "',"
        ZSql = ZSql + "'" + WValor29 + "',"
        ZSql = ZSql + "'" + WValor30 + "',"
        ZSql = ZSql + "'" + WValorNumero1 + "',"
        ZSql = ZSql + "'" + WValorNumero2 + "',"
        ZSql = ZSql + "'" + WValorNumero3 + "',"
        ZSql = ZSql + "'" + WValorNumero4 + "',"
        ZSql = ZSql + "'" + WValorNumero5 + "',"
        ZSql = ZSql + "'" + WValorNumero6 + "',"
        ZSql = ZSql + "'" + WValorNumero7 + "',"
        ZSql = ZSql + "'" + WValorNumero8 + "',"
        ZSql = ZSql + "'" + WValorNumero9 + "',"
        ZSql = ZSql + "'" + WValorNumero10 + "',"
        ZSql = ZSql + "'" + WValorNumero11 + "',"
        ZSql = ZSql + "'" + WValorNumero12 + "',"
        ZSql = ZSql + "'" + WValorNumero13 + "',"
        ZSql = ZSql + "'" + WValorNumero14 + "',"
        ZSql = ZSql + "'" + WValorNumero15 + "',"
        ZSql = ZSql + "'" + WValorNumero16 + "',"
        ZSql = ZSql + "'" + WValorNumero17 + "',"
        ZSql = ZSql + "'" + WValorNumero18 + "',"
        ZSql = ZSql + "'" + WValorNumero19 + "',"
        ZSql = ZSql + "'" + WValorNumero20 + "',"
        ZSql = ZSql + "'" + WValorNumero21 + "',"
        ZSql = ZSql + "'" + WValorNumero22 + "',"
        ZSql = ZSql + "'" + WValorNumero23 + "',"
        ZSql = ZSql + "'" + WValorNumero24 + "',"
        ZSql = ZSql + "'" + WValorNumero25 + "',"
        ZSql = ZSql + "'" + WValorNumero26 + "',"
        ZSql = ZSql + "'" + WValorNumero27 + "',"
        ZSql = ZSql + "'" + WValorNumero28 + "',"
        ZSql = ZSql + "'" + WValorNumero29 + "',"
        ZSql = ZSql + "'" + WValorNumero30 + "',"
        ZSql = ZSql + "'" + WEnsayo + "',"
        ZSql = ZSql + "'" + WAspecto + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WObserva2 + "',"
        ZSql = ZSql + "'" + WConfecciono + "',"
        ZSql = ZSql + "'" + WLiberada + "',"
        ZSql = ZSql + "'" + WDevuelta + "',"
        ZSql = ZSql + "'" + WLote + "',"
        ZSql = ZSql + "'" + WRechazo + "',"
        ZSql = ZSql + "'" + WNuevo + "',"
        ZSql = ZSql + "'" + WFechaord + "',"
        ZSql = ZSql + "'" + WDate + "')"
        
        spPrueart = ZSql
        Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
                        
    If Val(Devuelta.Text) > 0 Then
    
        Auxi = "2"
    
        WLote = NroRechazo.Text
        Call Ceros(WLote, 6)
        WPrueba = Auxi + WLote
        WProducto = Producto.Text
        WFecha = fecha.Text
        WOrden = Orden.Text
        
        WValor1 = Valor1.Text
        WValor2 = Valor2.Text
        WValor3 = Valor3.Text
        WValor4 = Valor4.Text
        WValor5 = Valor5.Text
        WValor6 = Valor6.Text
        WValor7 = Valor7.Text
        WValor8 = Valor8.Text
        WValor9 = Valor9.Text
        WValor10 = Valor10.Text
        WValor11 = Valor11.Text
        WValor12 = Valor12.Text
        WValor13 = Valor13.Text
        WValor14 = Valor14.Text
        WValor15 = Valor15.Text
        WValor16 = Valor16.Text
        WValor17 = Valor17.Text
        WValor18 = Valor18.Text
        WValor19 = Valor19.Text
        WValor20 = Valor20.Text
        WValor21 = Valor21.Text
        WValor22 = Valor22.Text
        WValor23 = Valor23.Text
        WValor24 = Valor24.Text
        WValor25 = Valor25.Text
        WValor26 = Valor26.Text
        WValor27 = Valor27.Text
        WValor28 = Valor28.Text
        WValor29 = Valor29.Text
        WValor30 = Valor30.Text
        
        WValorNumero1 = ValorNumero1.Text
        WValorNumero2 = ValorNumero2.Text
        WValorNumero3 = ValorNumero3.Text
        WValorNumero4 = ValorNumero4.Text
        WValorNumero5 = ValorNumero5.Text
        WValorNumero6 = ValorNumero6.Text
        WValorNumero7 = ValorNumero7.Text
        WValorNumero8 = ValorNumero8.Text
        WValorNumero9 = ValorNumero9.Text
        WValorNumero10 = ValorNumero10.Text
        WValorNumero11 = ValorNumero11.Text
        WValorNumero12 = ValorNumero12.Text
        WValorNumero13 = ValorNumero13.Text
        WValorNumero14 = ValorNumero14.Text
        WValorNumero15 = ValorNumero15.Text
        WValorNumero16 = ValorNumero16.Text
        WValorNumero17 = ValorNumero17.Text
        WValorNumero18 = ValorNumero18.Text
        WValorNumero19 = ValorNumero19.Text
        WValorNumero20 = ValorNumero20.Text
        WValorNumero21 = ValorNumero21.Text
        WValorNumero22 = ValorNumero22.Text
        WValorNumero23 = ValorNumero23.Text
        WValorNumero24 = ValorNumero24.Text
        WValorNumero25 = ValorNumero25.Text
        WValorNumero26 = ValorNumero26.Text
        WValorNumero27 = ValorNumero27.Text
        WValorNumero28 = ValorNumero28.Text
        WValorNumero29 = ValorNumero29.Text
        WValorNumero30 = ValorNumero30.Text
        
        WEnsayo = Ensayo.Text
        WAspecto = Aspecto.Text
        WObservaciones = Observaciones.Text
        WConfecciono = Confecciono.Text
        WLiberada = "0"
        WDevuelta = Devuelta.Text
        WLote = NroRechazo.Text
        WRechazo = NroRechazo.Text
        WNueva = Nueva.Text
        WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
        WDate = Date$
        WObserva2 = ""
    
        Rem XParam = "'" + WPrueba + "','" _
        rem         + WProducto + "','" _
        rem         + WFecha + "','" _
        rem         + WOrden + "','" _
        rem         + WValor1 + "','" _
        rem         + WValor2 + "','" _
        rem         + WValor3 + "','" _
        rem         + WValor4 + "','" _
        rem         + WValor5 + "','" _
        rem         + WValor6 + "','" _
        rem         + WValor7 + "','" _
        rem         + WValor8 + "','" _
        rem         + WValor9 + "','" _
        rem         + WValor10 + "','" _
        rem         + WEnsayo + "','" _
        rem         + WAspecto + "','" _
        rem         + WObservaciones + "','" _
        rem         + WObserva2 + "','" _
        rem         + WConfecciono + "','" _
        rem         + WLiberada + "','" _
        rem         + WDevuelta + "','" _
        rem         + WLote + "','" _
        rem         + WRechazo + "','" _
        rem         + WNueva + "','" + WFechaOrd + "','" _
        rem         + WDate + "'"
        Rem Set rstPrueart = db.OpenRecordset("AltaPrueart " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO PrueArt ("
        ZSql = ZSql + "Prueba ,"
        ZSql = ZSql + "Producto ,"
        ZSql = ZSql + "Fecha ,"
        ZSql = ZSql + "Orden ,"
        ZSql = ZSql + "Valor1 ,"
        ZSql = ZSql + "Valor2 ,"
        ZSql = ZSql + "Valor3 ,"
        ZSql = ZSql + "Valor4 ,"
        ZSql = ZSql + "Valor5 ,"
        ZSql = ZSql + "Valor6 ,"
        ZSql = ZSql + "Valor7 ,"
        ZSql = ZSql + "Valor8 ,"
        ZSql = ZSql + "Valor9 ,"
        ZSql = ZSql + "Valor10 ,"
        ZSql = ZSql + "Valor11 ,"
        ZSql = ZSql + "Valor12 ,"
        ZSql = ZSql + "Valor13 ,"
        ZSql = ZSql + "Valor14 ,"
        ZSql = ZSql + "Valor15 ,"
        ZSql = ZSql + "Valor16 ,"
        ZSql = ZSql + "Valor17 ,"
        ZSql = ZSql + "Valor18 ,"
        ZSql = ZSql + "Valor19 ,"
        ZSql = ZSql + "Valor20 ,"
        ZSql = ZSql + "Valor21 ,"
        ZSql = ZSql + "Valor22 ,"
        ZSql = ZSql + "Valor23 ,"
        ZSql = ZSql + "Valor24 ,"
        ZSql = ZSql + "Valor25 ,"
        ZSql = ZSql + "Valor26 ,"
        ZSql = ZSql + "Valor27 ,"
        ZSql = ZSql + "Valor28 ,"
        ZSql = ZSql + "Valor29 ,"
        ZSql = ZSql + "Valor30 ,"
        ZSql = ZSql + "ValorNumero1 ,"
        ZSql = ZSql + "ValorNumero2 ,"
        ZSql = ZSql + "ValorNumero3 ,"
        ZSql = ZSql + "ValorNumero4 ,"
        ZSql = ZSql + "ValorNumero5 ,"
        ZSql = ZSql + "ValorNumero6 ,"
        ZSql = ZSql + "ValorNumero7 ,"
        ZSql = ZSql + "ValorNumero8 ,"
        ZSql = ZSql + "ValorNumero9 ,"
        ZSql = ZSql + "ValorNumero10 ,"
        ZSql = ZSql + "ValorNumero11 ,"
        ZSql = ZSql + "ValorNumero12 ,"
        ZSql = ZSql + "ValorNumero13 ,"
        ZSql = ZSql + "ValorNumero14 ,"
        ZSql = ZSql + "ValorNumero15 ,"
        ZSql = ZSql + "ValorNumero16 ,"
        ZSql = ZSql + "ValorNumero17 ,"
        ZSql = ZSql + "ValorNumero18 ,"
        ZSql = ZSql + "ValorNumero19 ,"
        ZSql = ZSql + "ValorNumero20 ,"
        ZSql = ZSql + "ValorNumero21 ,"
        ZSql = ZSql + "ValorNumero22 ,"
        ZSql = ZSql + "ValorNumero23 ,"
        ZSql = ZSql + "ValorNumero24 ,"
        ZSql = ZSql + "ValorNumero25 ,"
        ZSql = ZSql + "ValorNumero26 ,"
        ZSql = ZSql + "ValorNumero27 ,"
        ZSql = ZSql + "ValorNumero28 ,"
        ZSql = ZSql + "ValorNumero29 ,"
        ZSql = ZSql + "ValorNumero30 ,"
        ZSql = ZSql + "Ensayo ,"
        ZSql = ZSql + "Aspecto ,"
        ZSql = ZSql + "Observaciones ,"
        ZSql = ZSql + "Observa2 ,"
        ZSql = ZSql + "Confecciono ,"
        ZSql = ZSql + "Liberada ,"
        ZSql = ZSql + "Devuelta ,"
        ZSql = ZSql + "Lote ,"
        ZSql = ZSql + "Rechazo ,"
        ZSql = ZSql + "Nueva ,"
        ZSql = ZSql + "FechaOrd ,"
        ZSql = ZSql + "WDate )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + WPrueba + "',"
        ZSql = ZSql + "'" + WProducto + "',"
        ZSql = ZSql + "'" + WFecha + "',"
        ZSql = ZSql + "'" + WOrden + "',"
        ZSql = ZSql + "'" + WValor1 + "',"
        ZSql = ZSql + "'" + WValor2 + "',"
        ZSql = ZSql + "'" + WValor3 + "',"
        ZSql = ZSql + "'" + WValor4 + "',"
        ZSql = ZSql + "'" + WValor5 + "',"
        ZSql = ZSql + "'" + WValor6 + "',"
        ZSql = ZSql + "'" + WValor7 + "',"
        ZSql = ZSql + "'" + WValor8 + "',"
        ZSql = ZSql + "'" + WValor9 + "',"
        ZSql = ZSql + "'" + WValor10 + "',"
        ZSql = ZSql + "'" + WValor11 + "',"
        ZSql = ZSql + "'" + WValor12 + "',"
        ZSql = ZSql + "'" + WValor13 + "',"
        ZSql = ZSql + "'" + WValor14 + "',"
        ZSql = ZSql + "'" + WValor15 + "',"
        ZSql = ZSql + "'" + WValor16 + "',"
        ZSql = ZSql + "'" + WValor17 + "',"
        ZSql = ZSql + "'" + WValor18 + "',"
        ZSql = ZSql + "'" + WValor19 + "',"
        ZSql = ZSql + "'" + WValor20 + "',"
        ZSql = ZSql + "'" + WValor21 + "',"
        ZSql = ZSql + "'" + WValor22 + "',"
        ZSql = ZSql + "'" + WValor23 + "',"
        ZSql = ZSql + "'" + WValor24 + "',"
        ZSql = ZSql + "'" + WValor25 + "',"
        ZSql = ZSql + "'" + WValor26 + "',"
        ZSql = ZSql + "'" + WValor27 + "',"
        ZSql = ZSql + "'" + WValor28 + "',"
        ZSql = ZSql + "'" + WValor29 + "',"
        ZSql = ZSql + "'" + WValor30 + "',"
        ZSql = ZSql + "'" + WValorNumero1 + "',"
        ZSql = ZSql + "'" + WValorNumero2 + "',"
        ZSql = ZSql + "'" + WValorNumero3 + "',"
        ZSql = ZSql + "'" + WValorNumero4 + "',"
        ZSql = ZSql + "'" + WValorNumero5 + "',"
        ZSql = ZSql + "'" + WValorNumero6 + "',"
        ZSql = ZSql + "'" + WValorNumero7 + "',"
        ZSql = ZSql + "'" + WValorNumero8 + "',"
        ZSql = ZSql + "'" + WValorNumero9 + "',"
        ZSql = ZSql + "'" + WValorNumero10 + "',"
        ZSql = ZSql + "'" + WValorNumero11 + "',"
        ZSql = ZSql + "'" + WValorNumero12 + "',"
        ZSql = ZSql + "'" + WValorNumero13 + "',"
        ZSql = ZSql + "'" + WValorNumero14 + "',"
        ZSql = ZSql + "'" + WValorNumero15 + "',"
        ZSql = ZSql + "'" + WValorNumero16 + "',"
        ZSql = ZSql + "'" + WValorNumero17 + "',"
        ZSql = ZSql + "'" + WValorNumero18 + "',"
        ZSql = ZSql + "'" + WValorNumero19 + "',"
        ZSql = ZSql + "'" + WValorNumero20 + "',"
        ZSql = ZSql + "'" + WValorNumero21 + "',"
        ZSql = ZSql + "'" + WValorNumero22 + "',"
        ZSql = ZSql + "'" + WValorNumero23 + "',"
        ZSql = ZSql + "'" + WValorNumero24 + "',"
        ZSql = ZSql + "'" + WValorNumero25 + "',"
        ZSql = ZSql + "'" + WValorNumero26 + "',"
        ZSql = ZSql + "'" + WValorNumero27 + "',"
        ZSql = ZSql + "'" + WValorNumero28 + "',"
        ZSql = ZSql + "'" + WValorNumero29 + "',"
        ZSql = ZSql + "'" + WValorNumero30 + "',"
        ZSql = ZSql + "'" + WEnsayo + "',"
        ZSql = ZSql + "'" + WAspecto + "',"
        ZSql = ZSql + "'" + WObservaciones + "',"
        ZSql = ZSql + "'" + WObserva2 + "',"
        ZSql = ZSql + "'" + WConfecciono + "',"
        ZSql = ZSql + "'" + WLiberada + "',"
        ZSql = ZSql + "'" + WDevuelta + "',"
        ZSql = ZSql + "'" + WLote + "',"
        ZSql = ZSql + "'" + WRechazo + "',"
        ZSql = ZSql + "'" + WNuevo + "',"
        ZSql = ZSql + "'" + WFechaord + "',"
        ZSql = ZSql + "'" + WDate + "')"
        
        spPrueart = ZSql
        Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    If Val(Liberada.Text) > 0 Then
    
        ZEntra = "N"
        If Left$(Producto.Text, 2) = "DY" Then
            XParam = "'" + Trim(PartidaProveedor.Text) + "','" _
                     + Producto.Text + "'"
            spLaudo = "ListaLaudoArticuloPartiOri " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                ZEntra = "S"
                ZZLaudoOriginal = rstLaudo!Laudo
                rstLaudo.Close
            End If
        End If
    
        WLaudo = Lote.Text
        WRenglon = "1"
        WFecha = fecha.Text
        WOrden = Orden.Text
        WArticulo = Producto.Text
        WLiberada = Liberada.Text
        WDevuelta = "0"
        WLote = Lote.Text
        WRechazo = ""
        WActualiza = "N"
        WMarca = ""
        WInforme = WInforme
        If ZEntra = "N" Then
            WSaldo = Liberada.Text
                Else
            WSaldo = "0"
        End If
        WOrigen = OrigenMercaderia.Text
        WPartiOri = Trim(PartidaProveedor.Text)
        WEnvase = Str$(XEnvase)
            
        Auxi1 = Str$(WLaudo)
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(WRenglon)
        Call Ceros(Auxi2, 2)
            
        WClave = Auxi1 + Auxi2
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                + WLaudo + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WArticulo + "','" _
                + WLiberada + "','" _
                + WDevuelta + "','" _
                + WOrden + "','" _
                + WMarca + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WInforme + "','" _
                + WActualiza + "','" _
                + WDate + "','" _
                + WSaldo + "','" _
                + WOrigen + "','" _
                + WPartiOri + "','" _
                + WEnvase + "'"
                
        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
                     
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "NroDespacho = " + "'" + ZNroDespacho + "',"
        ZSql = ZSql + "Procedencia = " + "'" + ZProcedencia + "',"
        ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
    
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        Rem dada
        Rem dada
        Rem dada
        Rem dada
        Rem dada
        
        If dada = 9999 Then
        
            Producto.Text = UCase(Producto.Text)
            spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                ZZTipoMp = IIf(IsNull(rstArticulo!TipoMp), "0", rstArticulo!TipoMp)
                rstArticulo.Close
                
                If ZZTipoMp = 1 Then
                
                    XEmpresa = WEmpresa
                    Select Case Val(WEmpresa)
                        Case 1, 3, 5, 6, 7, 10, 11
                            WEmpresa = "0003"
                            txtOdbc = "Empresa03"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        Case Else
                            WEmpresa = "0004"
                            
                            txtOdbc = "Empresa04"
                            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    End Select
            
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM Homologa"
                    ZSql = ZSql + " Where Proveedor = " + "'" + Proveedor.Text + "'"
                    ZSql = ZSql + " and CodigoMp = " + "'" + Producto.Text + "'"
                    ZSql = ZSql + " and Estado = " + "'" + "1" + "'"
                    spHomologa = ZSql
                    Set rstHomologa = db.OpenRecordset(spHomologa, dbOpenSnapshot, dbSQLPassThrough)
                    If rstHomologa.RecordCount > 0 Then
                        ZZIngre = "S"
                        rstHomologa.Close
                    End If
                    
                    Call Conecta_Empresa
                    
                    If ZZIngre = "N" Then
            
                        ZSql = ""
                        ZSql = ZSql + "Select *"
                        ZSql = ZSql + " FROM Orden"
                        ZSql = ZSql + " Where Articulo = " + "'" + WArticulo.Text + "'"
                        ZSql = ZSql + " and Proveedor = " + "'" + Proveedor.Text + "'"
                        spOrden = ZSql
                        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                        If rstOrden.RecordCount > 0 Then
                            ZZIngre = "S"
                            rstOrden.Close
                        End If
                        
                    End If
                    
                        Else
                        
                    ZZIngre = "S"
                    
                End If
                    
                If ZZIngre = "N" Then
                
                    If TipoOrden.ListIndex = 1 Then
                    
                        T$ = "Ingreso de Orden de Compra"
                        m$ = "Materia Prima homologable y no existe muestra aceptada. Desea continuar"
                        Respuesta% = MsgBox(m$, 32 + 4, T$)
                        If Respuesta% = 6 Then
                            m$ = "Coloque en homologacion los codigos de Materia Prima a Homologar"
                            G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                            ZZIngre = "S"
                                Else
                            Exit Sub
                        End If
                        
                            Else
                            
                        m$ = "Materia Prima homologable y no existe muestra aceptada"
                        G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                        Exit Sub
                        
                    End If
                    
                End If
                
                    Else
                    
                m$ = "Materia Prima Inexistentre"
                G% = MsgBox(m$, 0, "Ingreso de Orden de Compra")
                Exit Sub
                
            End If
            
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        If ZEntra = "S" Then
            XParam = "'" + Str$(ZZLaudoOriginal) + "','" _
                    + Producto.Text + "'"
            spLaudo = "ListaLaudoArticulo " + XParam
            Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            If rstLaudo.RecordCount > 0 Then
                WClave = rstLaudo!Clave
                WSaldo = Str$(rstLaudo!Saldo + Val(Liberada.Text))
                WDate = Date$
                rstLaudo.Close
            
                XParam = "'" + WClave + "','" _
                    + WDate + "','" _
                    + WSaldo + "'"
                spLaudo = "ModificaLaudoSaldo " + XParam
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
        
        
    End If
    
    If Val(Devuelta.Text) > 0 Then
    
        WLaudo = NroRechazo.Text
        WRenglon = "1"
        WFecha = fecha.Text
        WOrden = Orden.Text
        WArticulo = Producto.Text
        WLiberada = ""
        WDevuelta = Devuelta.Text
        WLote = NroRechazo.Text
        WRechazo = NroRechazo.Text
        WActualiza = Nueva.Text
        WMarca = ""
        WInforme = WInforme
        WSaldo = "0"
        WOrigen = OrigenMercaderia.Text
        WPartiOri = Trim(PartidaProveedor.Text)
        WEnvase = Str$(XEnvase)
            
        Auxi1 = Str$(WLaudo)
        Call Ceros(Auxi1, 6)
        Auxi2 = Str$(WRenglon)
        Call Ceros(Auxi2, 2)
            
        WClave = Auxi1 + Auxi2
        WDate = Date$
        
        XParam = "'" + WClave + "','" _
                + WLaudo + "','" _
                + WRenglon + "','" _
                + WFecha + "','" _
                + WArticulo + "','" _
                + WLiberada + "','" _
                + WDevuelta + "','" _
                + WOrden + "','" _
                + WMarca + "','" _
                + WLote + "','" _
                + WRechazo + "','" _
                + WInforme + "','" _
                + WActualiza + "','" _
                + WDate + "','" _
                + WSaldo + "','" _
                + WOrigen + "','" _
                + WPartiOri + "','" _
                + WEnvase + "'"
                
        Set rstLaudo = db.OpenRecordset("AltaLaudo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        WFechaord = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WLaudo + "','" _
                     + WFechaord + "'"
                     
        Set rstLaudo = db.OpenRecordset("ModificaLaudoFechaOrd " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        
        ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
    
        ZSql = ""
        ZSql = ZSql + "UPDATE Laudo SET "
        ZSql = ZSql + "NroDespacho = " + "'" + ZNroDespacho + "',"
        ZSql = ZSql + "Procedencia = " + "'" + ZProcedencia + "',"
        ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
        ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
        ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
    
        spLaudo = ZSql
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    WPrecio = 0
    
    For WDa% = 1 To 40
        Auxi3 = Orden.Text
        Call Ceros(Auxi3, 6)
        Auxi1 = WDa%
        Call Ceros(Auxi1, 2)
        WClave = Auxi3 + Auxi1
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WMoneda = rstOrden!Moneda
            WTipoOrden = IIf(IsNull(rstOrden!Tipo), "0", rstOrden!Tipo)
            If Producto.Text = rstOrden!Articulo Then
                WPrecio = rstOrden!Precio
                WLiberada = Str$(rstOrden!Liberada + Val(Liberada.Text))
                WDevuelta = Str$(rstOrden!Devuelta + Val(Devuelta.Text))
                WFechaentrega = fecha.Text
                WDate = Date$
                rstOrden.Close
                XParam = "'" + WClave + "','" _
                    + WLiberada + "','" _
                    + WDevuelta + "','" _
                    + WFechaentrega + "','" _
                    + WDate + "'"
                spOrden = "ModificaOrdenPrueba " + XParam
                Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                rstOrden.Close
            End If
        End If
    Next WDa%
    
    spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        WProducto = Producto.Text
        If Nueva.Text = "S" Then
            WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Liberada) - Val(Devuelta))
                Else
            WLaboratorio = Str$(rstArticulo!Laboratorio - Val(Liberada))
        End If
        
        Select Case WTipoOrden
            Case 1, 2
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo1), "0", rstArticulo!Costo1)
                    XCosto3 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto1), "0", rstArticulo!WCosto1)
                    XCosto3 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                    WCosto1 = Str$(XCosto1)
                    WCosto3 = Str$(XCosto3)
                End If
                WEntradas = Str$(rstArticulo!Entradas + Val(Liberada))
                WDate = Date$
                rstArticulo.Close
            
            Case Else
                XStock1 = rstArticulo!Inicial + rstArticulo!Entradas - rstArticulo!Salidas
                If WMoneda = 0 Then
                    XCosto1 = IIf(IsNull(rstArticulo!Costo3), "0", rstArticulo!Costo3)
                        Else
                    XCosto1 = IIf(IsNull(rstArticulo!WCosto3), "0", rstArticulo!WCosto3)
                End If
                XCostoTotal1 = XStock1 * XCosto1
                
                XStock2 = Val(Liberada)
                XCosto2 = WPrecio
                XCostoTotal2 = XStock2 * XCosto2
                
                XCosto = 0
                XStock = XStock1 + XStock2
                XCostoTotal = XCostoTotal1 + XCostoTotal2
                If XStock <> 0 Then
                    XCosto = XCostoTotal / XStock
                End If
                
                Call Redondeo(XCosto)
                    
                WCosto1 = Str$(WPrecio)
                WCosto3 = Str$(XCosto)
                
                WEntradas = Str$(rstArticulo!Entradas + Val(Liberada))
                WDate = Date$
                rstArticulo.Close
                
                ZZEmpresa = WEmpresa
                
                If WMoneda = 0 Then
                    Rem U$S
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " ZCosto1 = " + "'" + WCosto1 + "',"
                    ZSql = ZSql + " OrdenIII = " + "'" + Orden.Text + "',"
                    ZSql = ZSql + " PtaOrdenIII = " + "'" + ZZEmpresa + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WProducto + "'"
                        Else
                    Rem $
                    ZSql = ""
                    ZSql = ZSql + "UPDATE Articulo SET "
                    ZSql = ZSql + " WCosto1 = " + "'" + WCosto1 + "',"
                    ZSql = ZSql + " OrdenII = " + "'" + Orden.Text + "',"
                    ZSql = ZSql + " PtaOrdenII = " + "'" + ZZEmpresa + "'"
                    ZSql = ZSql + " Where Codigo = " + "'" + WProducto + "'"
                End If
                
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        
                Rem actualiza los datos de la empresa
                
                XEmpresa = WEmpresa
                
                WCodigo = WProducto
                XParam = "'" + WCodigo + "','" _
                             + WCosto1 + "'"
                    
                WEmpresa = "0001"
                txtOdbc = "Empresa01"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                WEmpresa = "0002"
                txtOdbc = "Empresa02"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                WEmpresa = "0005"
                txtOdbc = "Empresa05"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                WEmpresa = "0006"
                txtOdbc = "Empresa06"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                
                WEmpresa = "0007"
                txtOdbc = "Empresa07"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                
                WEmpresa = "0008"
                txtOdbc = "Empresa08"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                
                WEmpresa = "0009"
                txtOdbc = "Empresa09"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                
                WEmpresa = "0010"
                txtOdbc = "Empresa10"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                
                WEmpresa = "0011"
                txtOdbc = "Empresa11"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    
                spArticulo = ZSql
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    
                    
                    
                    
                Select Case Val(XEmpresa)
                    Case 1
                        WEmpresa = "0001"
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 2
                        WEmpresa = "0002"
                        txtOdbc = "Empresa02"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 3
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 4
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 5
                        WEmpresa = "0005"
                        txtOdbc = "Empresa05"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 6
                        WEmpresa = "0006"
                        txtOdbc = "Empresa06"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 7
                        WEmpresa = "0007"
                        txtOdbc = "Empresa07"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 8
                        WEmpresa = "0008"
                        txtOdbc = "Empresa08"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 9
                        WEmpresa = "0009"
                        txtOdbc = "Empresa09"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 10
                        WEmpresa = "0010"
                        txtOdbc = "Empresa10"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case 11
                        WEmpresa = "0011"
                        txtOdbc = "Empresa11"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                End Select
                
        End Select
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Articulo SET "
        ZSql = ZSql + "Laboratorio = " + "'" + WLaboratorio + "',"
        ZSql = ZSql + "Entradas = " + "'" + WEntradas + "',"
        ZSql = ZSql + "WDate = " + "'" + WDate + "'"
        ZSql = ZSql + " Where Codigo = " + "'" + WProducto + "'"
        
        spArticulo = ZSql
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        
    End If
    
    If Val(Devuelta.Text) > 0 Then
    
        Sql1 = "Select * "
        Sql2 = " FROM Orden"
        Sql3 = " Where Orden = " + "'" + Orden.Text + "'"
        spOrden = Sql1 + Sql2 + Sql3
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            WProveedor = rstOrden!Proveedor
            rstOrden.Close
        End If
    
        spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
        Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            WDesProveedor = RstProveedor!Nombre
            RstProveedor.Close
        End If
    
        If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        
            WEmail = "pcorna@surfactan.com.ar"
            sTo = WEmail
            sCC = ""
            sBCC = ""
            sSubject = "Rechazo de M.P."
        
            sBody = "Orden = " + Orden.Text + " - " + _
            "Proveedor = " + RTrim(WDesProveedor) + " - " + _
            "Producto = " + Producto.Text + " " + RTrim(Descriprod.Caption) + " - " + _
            "Cantidad = " + Devuelta.Text
    
            ret = Shell("Start.exe " _
                    & "mailto:" & """" & sTo & """" _
                    & "?Subject=" & """" & sSubject & """" _
                    & "&cc=" & """" & sCC & """" _
                    & "&bcc=" & """" & sBCC & """" _
                    & "&Body=" & """" & sBody & """" _
                    & "&File=" & """" & "c:\autoexec.bat" & """" _
                    , 0)
                    
                Else
                
            WEmail = "Rechazo"
            sTo = WEmail
            sCC = ""
            sBCC = ""
            sSubject = "Rechazo de M.P."
            sBody = "Orden = " + Orden.Text + " - " + _
                    "Proveedor = " + RTrim(WDesProveedor) + " - " + _
                    "Producto = " + Producto.Text + " " + RTrim(Descriprod.Caption) + " - " + _
                    "Cantidad = " + Devuelta.Text
            SFile = ""
        
            EmailAddress = sTo
            CopiaAddress = sCC
            MSubject = sSubject
            MBody = sBody
            MAttach = ""
            MAttachI = ""
            MAttachII = ""
            MAttachIII = ""
            MAttachIV = ""
            MAttachVI = ""
            MAttachVII = ""
            MAttachVIII = ""
            
            SendEmail
                
        End If
                
    End If
   
    If Val(Liberada.Text) > 0 Then
        Entra = "N"
        Select Case Val(WEmpresa)
            Case 1
                If Val(Lote.Text) >= 190000 And Val(Lote.Text) <= 194999 Then
                    Entra = "S"
                End If
                If Val(Lote.Text) >= 990000 And Val(Lote.Text) <= 994999 Then
                    Entra = "S"
                End If
            Case 2
                If Val(Lote.Text) >= 690000 And Val(Lote.Text) <= 691999 Then
                    Entra = "S"
                End If
            Case 3
                If Val(Lote.Text) >= 290000 And Val(Lote.Text) <= 294999 Then
                    Entra = "S"
                End If
            Case 4
                If Val(Lote.Text) >= 790000 And Val(Lote.Text) <= 794999 Then
                    Entra = "S"
                End If
            Case 5
                If Val(Lote.Text) >= 390000 And Val(Lote.Text) <= 394999 Then
                    Entra = "S"
                End If
            Case 6
                If Val(Lote.Text) >= 490000 And Val(Lote.Text) <= 494999 Then
                    Entra = "S"
                End If
            Case 7
                If Val(Lote.Text) >= 590000 And Val(Lote.Text) <= 594999 Then
                    Entra = "S"
                End If
            Case 8
                If Val(Lote.Text) >= 890000 And Val(Lote.Text) <= 894999 Then
                    Entra = "S"
                End If
            Case 9
                If Val(Lote.Text) >= 692000 And Val(Lote.Text) <= 694999 Then
                    Entra = "S"
                End If
            Case Else
        End Select
        
        If Entra = "S" Then
        
            Sql1 = "Select * "
            Sql2 = " FROM Orden"
            Sql3 = " Where Orden = " + "'" + Orden.Text + "'"
            spOrden = Sql1 + Sql2 + Sql3
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                WProveedor = rstOrden!Proveedor
                rstOrden.Close
            End If
    
            spProveedor = "ConsultaProveedores " + "'" + WProveedor + "'"
            Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                WDesProveedor = RstProveedor!Nombre
                RstProveedor.Close
            End If
    
            If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                WEmail = "pcorna@surfactan.com.ar"
                    Else
                WEmail = "Rechazo"
            End If
            sTo = WEmail
            sCC = ""
            sBCC = ""
            sSubject = "Aprobacion por Desvio de M.P."
        
            sBody = "Orden = " + Orden.Text + " - " + _
            "Proveedor = " + RTrim(WDesProveedor) + " - " + _
            "Producto = " + Producto.Text + " " + RTrim(Descriprod.Caption) + " - " + _
            "Cantidad = " + Liberada.Text
    
            ret = Shell("Start.exe " _
                & "mailto:" & """" & sTo & """" _
                & "?Subject=" & """" & sSubject & """" _
                & "&cc=" & """" & sCC & """" _
                & "&bcc=" & """" & sBCC & """" _
                & "&Body=" & """" & sBody & """" _
                & "&File=" & """" & "c:\autoexec.bat" & """" _
                , 0)
        End If
        
    End If
    
    
    If Val(Liberada.Text) > 0 Then
        T$ = "Ingreso de Pruebas"
        m$ = "Desea Imprimir las Etiquetas Correspondientes"
        Respuesta% = MsgBox(m$, 32 + 4, T$)
        If Respuesta% = 6 Then
            ImpreEtiqueta.Height = 2055
            ImpreEtiqueta.Left = 2280
            ImpreEtiqueta.Top = 1440
            ImpreEtiqueta.Width = 5535
            Kilos.Text = ""
            Cantidad.Text = ""
            ImpreEtiqueta.Visible = True
            Kilos.SetFocus
            DoEvents
            Exit Sub
        End If
    End If
    
    
    Call CmdLimpiar_Click
    panLote.Visible = False
    Producto.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub CancelaEtiqueta_Click()
    ImpreEtiqueta.Visible = False
    Call CmdLimpiar_Click
    panLote.Visible = False
    Producto.SetFocus
End Sub

Private Sub AceptaEtiqueta_Click()

    On Error GoTo WError
    
    ImpreEtiqueta.Visible = False
    
    OPEN_FILE_Etiqueta
    
    Salida = "N"
    DA = 0
    With rstEtiqueta
        .Index = "Codigo"
        .Seek ">=", DA
        If .NoMatch = False Then
            Do
                m$ = "EL proceso de Imprsion de Etiquetas ya se encuentra en proceso de impresion desde otra estacion"
                G% = MsgBox(m$, 0, "Impresion de Etiquetas")
                Salida = "S"
                Exit Do
            Loop
        End If
    End With
    
    If Salida <> "S" Then
    
        WClase = ""
        WIntervencion = ""
        WNaciones = ""
        WEmbalaje = ""
        ZMeses = "0"
        
        spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZMeses = rstArticulo!Meses
            WClase = IIf(IsNull(rstArticulo!Clase), "", rstArticulo!Clase)
            WIntervencion = IIf(IsNull(rstArticulo!Intervencion), "", rstArticulo!Intervencion)
            WNaciones = IIf(IsNull(rstArticulo!Naciones), "", rstArticulo!Naciones)
            WEmbalaje = IIf(IsNull(rstArticulo!Embalaje), "", rstArticulo!Embalaje)
            rstArticulo.Close
        End If
        
        ZInforme = 0
        Sql1 = "Select * "
        Sql2 = " FROM Informe"
        Sql3 = " Where Orden = " + "'" + Orden.Text + "'"
        Sql4 = " and Articulo = " + "'" + Producto.Text + "'"
        spInforme = Sql1 + Sql2 + Sql3 + Sql4
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            ZInforme = rstInforme!Informe
            rstInforme.Close
        End If
    
        DA = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", DA
                If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
        WNeto = Val(Kilos.Text)
        
        ZCantidad = Int(Val(Cantidad.Text) / 2)
        If ZCantidad * 2 <> Val(Cantidad.Text) Then
            ZCantidad = ZCantidad + 1
        End If
        
        If Val(WEmpresa) <> 5 Then
        
            If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
            
                WMes = Val(Mid$(fecha.Text, 4, 2))
                WAno = Val(Right$(fecha.Text, 4))
                For ZCiclo = 1 To ZMeses
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                Next ZCiclo
            
                XMes = Str$(WMes)
                XAno = Str$(WAno)
                Call Ceros(XMes, 2)
                Call Ceros(XAno, 4)
                If Val(Left$(fecha.Text, 2)) <= 30 Then
                    If Val(XMes) = 2 And Val(Left$(fecha.Text, 2)) > 28 Then
                        ZVencimiento = "28/" + XMes + "/" + XAno
                            Else
                        ZVencimiento = Left$(fecha.Text, 3) + XMes + "/" + XAno
                    End If
                        Else
                    If Val(XMes) = 2 Then
                        ZVencimiento = "28/" + XMes + "/" + XAno
                            Else
                        ZVencimiento = "30/" + XMes + "/" + XAno
                    End If
                End If
            
            End If
        
            With rstEtiqueta
                For DA = 1 To ZCantidad
                    .Index = "Codigo"
                    .AddNew
                
                    ZLote = Lote.Text
                    Call Ceros(ZLote, 6)
                
                    ZDa = Int((DA - 1) / 2)
                
                    !Codigo = DA
                    !Terminado = Producto.Text
                    !Lote = ZLote
                    !Cliente = ""
                    !Cantidad = Val(Kilos.Text)
                    !Nombre = "Fec.Lau.: " + fecha.Text
                    If ZVencimiento <> "00/00/0000" Then
                        !Impre1 = "Fec.Rea.:" + ZVencimiento
                            Else
                        !Impre1 = ""
                    End If
                    !Conservacion = !Impre1
                    !razon = "L.: " + Lote.Text
                    !DirEntrega = Kilos.Text + " Kgs."
                    !Clase = WClase
                    !Intervencion = WIntervencion
                    !Naciones = WNaciones
                    !Embalaje = WEmbalaje
                    !Bruto = 0
                    !Neto = ZDa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                        !Observaciones = "CONTROL CALIDAD"
                            Else
                        !Observaciones = "C.C.   PELLITAL"
                    End If
                
                    .Update
                Next DA
            End With

            ListaII.WindowTitle = "Emision de Etiquetas"
            ListaII.WindowTop = 0
            ListaII.WindowLeft = 0
            ListaII.WindowWidth = Screen.Width
            ListaII.WindowHeight = Screen.Height

            Select Case Mid$(WClase, 1, 1)
                Case "3"
                    ListaII.ReportFileName = "WEtiVerde3.rpt"
                Case "4"
                    ListaII.ReportFileName = "WEtiVerde4.rpt"
                Case "5"
                    ListaII.ReportFileName = "WEtiVerde5.rpt"
                Case "6"
                    ListaII.ReportFileName = "WEtiVerde6.rpt"
                Case "8"
                    ListaII.ReportFileName = "WEtiVerde8.rpt"
                Case "9"
                    ListaII.ReportFileName = "WEtiVerde9.rpt"
                Case Else
                    ListaII.ReportFileName = "WEtiVerde.rpt"
            End Select
                
            Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
            Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
            Rem Listado.Connect = Connect()
    
            ListaII.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
            ListaII.Destination = 1
            ListaII.PrinterCopies = 1
            ListaII.Action = 1
            
            
                    Else
                    
                    
            If ZVencimiento = "  /  /    " Or ZVencimiento = "00/00/0000" Then
        
                WMes = Val(Mid$(fecha.Text, 4, 2))
                WAno = Val(Right$(fecha.Text, 4))
                For ZCiclo = 1 To ZMeses
                    WMes = WMes + 1
                    If WMes > 12 Then
                        WAno = WAno + 1
                        WMes = 1
                    End If
                Next ZCiclo
            
                XMes = Str$(WMes)
                XAno = Str$(WAno)
                Call Ceros(XMes, 2)
                Call Ceros(XAno, 4)
                If Val(Left$(fecha.Text, 2)) <= 30 Then
                    If Val(XMes) = 2 And Val(Left$(fecha.Text, 2)) > 28 Then
                        ZVencimiento = "28/" + XMes + "/" + XAno
                            Else
                        ZVencimiento = Left$(fecha.Text, 3) + XMes + "/" + XAno
                    End If
                        Else
                    If Val(XMes) = 2 Then
                        ZVencimiento = "28/" + XMes + "/" + XAno
                            Else
                        ZVencimiento = "30/" + XMes + "/" + XAno
                    End If
                End If
            
            End If
            
            With rstEtiqueta
                For DA = 1 To ZCantidad
                    .Index = "Codigo"
                    .AddNew
                
                    ZLote = Lote.Text
                    Call Ceros(ZLote, 6)
                
                    ZCantidad = Kilos.Text
                    Call Ceros(ZCantidad, 4)
                
                    ZDa = Int((DA - 1) / 2)
                
                    !Codigo = DA
                    !Terminado = Producto.Text
                    !Lote = ZLote
                    !Cliente = ""
                    !Cantidad = Val(Kilos.Text)
                    !Nombre = Left$(Descriprod.Caption, 30)
                    If ZVencimiento <> "00/00/0000" Then
                        !DirEntrega = "Fecha Reanalisis : " + ZVencimiento
                            Else
                        !DirEntrega = ""
                    End If
                    !razon = "Lote : " + Lote.Text
                    Rem !DirEntrega = "Cantidad por Bulto : " + Kilos.Text + " Kgs."
                    Rem !DirEntrega = ""
                    !Conservacion = "Fecha de Ingreso : " + fecha.Text
                    !Impre1 = "Informe Nro.:" + Trim(Str$(ZInforme))
                    !Clase = WClase
                    !Intervencion = WIntervencion
                    !Naciones = WNaciones
                    !Embalaje = WEmbalaje
                    !Bruto = 0
                    !Neto = ZDa
                    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
                        !Observaciones = "CONTROL CALIDAD"
                            Else
                        !Observaciones = "C.C.   PELLITAL"
                    End If
                    !ConservacionII = Descriprod.Caption
                    
                    .Update
                Next DA
            End With
    
            ListaII.WindowTitle = "Emision de Etiquetas"
            ListaII.WindowTop = 0
            ListaII.WindowLeft = 0
            ListaII.WindowWidth = Screen.Width
            ListaII.WindowHeight = Screen.Height
        
            Rem ListaII.ReportFileName = "WEtiVerdeFarma.rpt"
            Select Case Mid$(WClase, 1, 1)
                Case "3"
                    ListaII.ReportFileName = "WEtiVerdeFarma3.rpt"
                Case "5"
                    ListaII.ReportFileName = "WEtiVerdeFarma5.rpt"
                Case "6"
                    ListaII.ReportFileName = "WEtiVerdeFarma6.rpt"
                Case "8"
                    ListaII.ReportFileName = "WEtiVerdeFarma8.rpt"
                Case "9"
                    ListaII.ReportFileName = "WEtiVerdeFarma9.rpt"
                Case Else
                    ListaII.ReportFileName = "WEtiVerdeFarma.rpt"
            End Select


            Rem Listado.ReportFileName = "WEtiVerde.rpt"
            Rem Listado.GroupSelectionFormula = Uno + Dos + Tres + Cuatro
            Rem Listado.DataFiles(0) = WEmpresa + "vent.mdb"
            Rem Listado.Connect = Connect()
    
            ListaII.DataFiles(0) = WEmpresa + "Auxi.mdb"
    
            ListaII.Destination = 1
            ListaII.PrinterCopies = 1
            ListaII.Action = 1
            
        End If
            
    
        DA = 0
        With rstEtiqueta
            .Index = "Codigo"
            .Seek ">=", DA
            If .NoMatch = False Then
                Do
                    .Delete
                    .MoveNext
                    If .EOF = True Then
                        Exit Do
                    End If
                Loop
            End If
        End With
    
    End If
    
    Call CmdLimpiar_Click
    panLote.Visible = False
    Producto.SetFocus
    
    Exit Sub

WError:

    Resume Next
    
End Sub

Private Sub Kilos_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cantidad.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Kilos.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Confirmaestado_Click()

    If CertificadoSi.Value = 1 And CertificadoNo.Value = 1 Then
        m$ = "Datos del Certificado de Analisis incorrectos"
        d% = MsgBox(m$, 0, "Ingreso de Ensayos de Materia Prima")
        Exit Sub
    End If
    If CertificadoSi.Value = 0 And CertificadoNo.Value = 0 Then
        m$ = "Datos del Certificado de Analisis incorrectos"
        d% = MsgBox(m$, 0, "Ingreso de Ensayos de Materia Prima")
        Exit Sub
    End If
    If EstadoSi.Value = 1 And EstadoNo.Value = 1 Then
        m$ = "Datos del Estado de Envases incorrectos"
        d% = MsgBox(m$, 0, "Ingreso de Ensayos de Materia Prima")
        Exit Sub
    End If
    If EstadoSi.Value = 0 And EstadoNo.Value = 0 Then
        m$ = "Datos del Estado de Envases incorrectos"
        d% = MsgBox(m$, 0, "Ingreso de Ensayos de Materia Prima")
        Exit Sub
    End If
    
    If CertificadoNo.Value = 1 Then
        WCertificado1 = "0"
    End If
    If CertificadoSi.Value = 1 Then
        WCertificado1 = "1"
    End If
    
    If EstadoNo.Value = 1 Then
        WEstado1 = "0"
    End If
    If EstadoSi.Value = 1 Then
        WEstado1 = "1"
    End If
    
    If Vencimiento.Text <> "  /  /    " And Vencimiento.Text <> "00/00/0000" Then
        Call Valida_fecha(Vencimiento.Text, Auxi4)
        If Auxi4 <> "S" Then
            m$ = "Fecha de Vencimiento incorrecta"
            d% = MsgBox(m$, 0, "Ingreso de Ensayos de Materia Prima")
            Exit Sub
        End If
    End If
    
    ZVencimiento = Vencimiento.Text
    WClave = ""
    
    ZSql = ""
    ZSql = ZSql + "Select * "
    ZSql = ZSql + " FROM Informe"
    ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
    ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
    ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        WClave = rstInforme!Clave
        rstInforme.Close
    End If
    
    ZVencimiento = Vencimiento.Text
    ZOrdVencimiento = Right$(ZVencimiento, 4) + Mid$(ZVencimiento, 4, 2) + Left$(ZVencimiento, 2)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Informe SET "
    ZSql = ZSql + "Certificado1 = " + "'" + WCertificado1 + "',"
    ZSql = ZSql + "Certificado2 = " + "'" + Certificado2.Text + "',"
    ZSql = ZSql + "Estado1 = " + "'" + WEstado1 + "',"
    ZSql = ZSql + "Estado2 = " + "'" + Estado2.Text + "',"
    ZSql = ZSql + "FechaVencimiento = " + "'" + ZVencimiento + "',"
    ZSql = ZSql + "OrdFechaVencimiento = " + "'" + ZOrdVencimiento + "'"
    ZSql = ZSql + " Where Clave = " + "'" + WClave + "'"
    
    spInforme = ZSql
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    
    IngresaEstado.Visible = False
    ValorNumero1.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    Producto.Text = "  -   -   "
    fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Orden.Text = ""
    
    Rem Ensayo1.Caption = ""
    Valor1.Text = ""
    ValorNumero1.Text = ""
    Rem Ensayo2.Caption = ""
    Valor2.Text = ""
    ValorNumero2.Text = ""
    Rem Ensayo3.Caption = ""
    Valor3.Text = ""
    ValorNumero3.Text = ""
    Rem Ensayo4.Caption = ""
    Valor4.Text = ""
    ValorNumero4.Text = ""
    Rem Ensayo5.Caption = ""
    Valor5.Text = ""
    ValorNumero5.Text = ""
    Rem Ensayo6.Caption = ""
    Valor6.Text = ""
    ValorNumero6.Text = ""
    Rem Ensayo7.Caption = ""
    Valor7.Text = ""
    ValorNumero7.Text = ""
    Rem Ensayo8.Caption = ""
    Valor8.Text = ""
    ValorNumero8.Text = ""
    Rem Ensayo9.Caption = ""
    Valor9.Text = ""
    ValorNumero9.Text = ""
    Rem Ensayo10.Caption = ""
    Valor10.Text = ""
    ValorNumero10.Text = ""
    Rem Ensayo11.Caption = ""
    Valor11.Text = ""
    ValorNumero11.Text = ""
    Rem Ensayo12.Caption = ""
    Valor12.Text = ""
    ValorNumero12.Text = ""
    Rem Ensayo13.Caption = ""
    Valor13.Text = ""
    ValorNumero13.Text = ""
    Rem Ensayo14.Caption = ""
    Valor14.Text = ""
    ValorNumero14.Text = ""
    Rem Ensayo15.Caption = ""
    Valor15.Text = ""
    ValorNumero15.Text = ""
    Rem Ensayo16.Caption = ""
    Valor16.Text = ""
    ValorNumero16.Text = ""
    Rem Ensayo17.Caption = ""
    Valor17.Text = ""
    ValorNumero17.Text = ""
    Rem Ensayo18.Caption = ""
    Valor18.Text = ""
    ValorNumero18.Text = ""
    Rem Ensayo19.Caption = ""
    Valor19.Text = ""
    ValorNumero19.Text = ""
    Rem Ensayo20.Caption = ""
    Valor20.Text = ""
    ValorNumero20.Text = ""
    
    Valor21.Text = ""
    ValorNumero21.Text = ""
    Valor22.Text = ""
    ValorNumero22.Text = ""
    Valor23.Text = ""
    ValorNumero23.Text = ""
    Valor24.Text = ""
    ValorNumero24.Text = ""
    Valor25.Text = ""
    ValorNumero25.Text = ""
    Valor26.Text = ""
    ValorNumero26.Text = ""
    Valor27.Text = ""
    ValorNumero27.Text = ""
    Valor28.Text = ""
    ValorNumero28.Text = ""
    Valor29.Text = ""
    ValorNumero29.Text = ""
    Valor30.Text = ""
    ValorNumero30.Text = ""
    
    ZEnsayo1 = ""
    ZEnsayo2 = ""
    ZEnsayo3 = ""
    ZEnsayo4 = ""
    ZEnsayo5 = ""
    ZEnsayo6 = ""
    ZEnsayo7 = ""
    ZEnsayo8 = ""
    ZEnsayo9 = ""
    ZEnsayo10 = ""
    ZEnsayo11 = ""
    ZEnsayo12 = ""
    ZEnsayo13 = ""
    ZEnsayo14 = ""
    ZEnsayo15 = ""
    ZEnsayo16 = ""
    ZEnsayo17 = ""
    ZEnsayo18 = ""
    ZEnsayo19 = ""
    ZEnsayo20 = ""
    ZEnsayo21 = ""
    ZEnsayo22 = ""
    ZEnsayo23 = ""
    ZEnsayo24 = ""
    ZEnsayo25 = ""
    ZEnsayo26 = ""
    ZEnsayo27 = ""
    ZEnsayo28 = ""
    ZEnsayo29 = ""
    ZEnsayo30 = ""
    
    Descriprod.Caption = ""
    
    Rem Descri1.Caption = ""
    Rem descri2.Caption = ""
    Rem Descri3.Caption = ""
    Rem Descri4.Caption = ""
    Rem Descri5.Caption = ""
    Rem Descri6.Caption = ""
    Rem Descri7.Caption = ""
    Rem Descri8.Caption = ""
    Rem Descri9.Caption = ""
    Rem Descri10.Caption = ""
    Rem Descri11.Caption = ""
    Rem Descri12.Caption = ""
    Rem Descri13.Caption = ""
    Rem Descri14.Caption = ""
    Rem Descri15.Caption = ""
    Rem Descri16.Caption = ""
    Rem Descri17.Caption = ""
    Rem Descri18.Caption = ""
    Rem Descri19.Caption = ""
    Rem Descri20.Caption = ""
    
    Ensayo.Text = ""
    Aspecto.Text = ""
    Observaciones.Text = ""
    Confecciono.Text = ""
    
    Std1.Caption = ""
    Std2.Caption = ""
    Std3.Caption = ""
    Std4.Caption = ""
    Std5.Caption = ""
    Std6.Caption = ""
    Std7.Caption = ""
    Std8.Caption = ""
    Std9.Caption = ""
    Std10.Caption = ""
    Std11.Caption = ""
    Std12.Caption = ""
    Std13.Caption = ""
    Std14.Caption = ""
    Std15.Caption = ""
    Std16.Caption = ""
    Std17.Caption = ""
    Std18.Caption = ""
    Std19.Caption = ""
    Std20.Caption = ""
    Std21.Caption = ""
    Std22.Caption = ""
    Std23.Caption = ""
    Std24.Caption = ""
    Std25.Caption = ""
    Std26.Caption = ""
    Std27.Caption = ""
    Std28.Caption = ""
    Std29.Caption = ""
    Std30.Caption = ""
    
    Partida.Text = ""
    OrigenMercaderiaII.Text = ""
    PartidaProveedorII.Text = ""
    NroRevalida.Text = ""
    RevalidaAnterior.Text = ""
    Vto.Text = "  /  /    "
    lblresultado.Caption = "Valor Standard"
    lblresultadoII.Caption = "Valor Standard"
    Producto.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgPruart.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(fecha.Text, Auxi4)
        If Auxi4 = "S" Then
            Informe.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Impensayo_Click()

    If Val(Auxi) = 0 Then
        Auxi = "0"
    End If
    
    If Val(Lote.Text) = 0 Then
        Lote.Text = "000000"
    End If

    Rem lista.ReportFileName = "Ensayo.rpt"
    Rem lista.GroupSelectionFormula = "{Prueart.Prueba} = " + Chr$(34) + Auxi + Lote.Text + Chr$(34)
    Rem lista.Destination = 1
    Rem lista.Action = 1
    
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            WAuxiliar = !Nombre
        End If
    End With
    
    Printer.Font = "Times New Roman"
    Printer.FontSize = "12"
    Printer.Print Tab(1); ""
    Printer.FontSize = "10"
    
    Printer.Print Tab(1); "Empresa : " + WAuxiliar
    Printer.Print Tab(1); ""
    Printer.Print Tab(20); "ENSAYO DE MATERIA PRIMA"
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Prueba"; Tab(15); Lote.Text
    Printer.Print Tab(1); "Producto"; Tab(15); Producto.Text; Tab(40); Descriprod.Caption
    Printer.Print Tab(1); "Fecha"; Tab(15); fecha.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(1); Tab(25); Descri1.Caption; Tab(80); Std1.Caption; Tab(105); Valor1.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(2); Tab(25); Descri2.Caption; Tab(80); Std2.Caption; Tab(105); Valor2.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(3); Tab(25); Descri3.Caption; Tab(80); Std3.Caption; Tab(105); Valor3.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(4); Tab(25); Descri4.Caption; Tab(80); Std4.Caption; Tab(105); Valor4.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(5); Tab(25); Descri5.Caption; Tab(80); Std5.Caption; Tab(105); Valor5.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(6); Tab(25); Descri6.Caption; Tab(80); Std6.Caption; Tab(105); Valor6.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(7); Tab(25); Descri7.Caption; Tab(80); Std7.Caption; Tab(105); Valor7.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(8); Tab(25); Descri8.Caption; Tab(80); Std8.Caption; Tab(105); Valor8.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(9); Tab(25); Descri9.Caption; Tab(80); Std9.Caption; Tab(105); Valor9.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(10); Tab(25); Descri10.Caption; Tab(80); Std10.Caption; Tab(105); Valor10.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(11); Tab(25); Descri11.Caption; Tab(80); Left$(Std11.Caption, 20); Tab(105); Valor11.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(12); Tab(25); Descri12.Caption; Tab(80); Left$(Std12.Caption, 20); Tab(105); Valor12.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(13); Tab(25); Descri13.Caption; Tab(80); Left$(Std13.Caption, 20); Tab(105); Valor13.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(14); Tab(25); Descri14.Caption; Tab(80); Left$(Std14.Caption, 20); Tab(105); Valor14.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(15); Tab(25); Descri15.Caption; Tab(80); Left$(Std15.Caption, 20); Tab(105); Valor15.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(16); Tab(25); Descri16.Caption; Tab(80); Left$(Std16.Caption, 20); Tab(105); Valor16.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(17); Tab(25); Descri17.Caption; Tab(80); Left$(Std17.Caption, 20); Tab(105); Valor17.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(18); Tab(25); Descri18.Caption; Tab(80); Left$(Std18.Caption, 20); Tab(105); Valor18.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(19); Tab(25); Descri19.Caption; Tab(80); Left$(Std19.Caption, 20); Tab(105); Valor19.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(20); Tab(25); Descri20.Caption; Tab(80); Left$(Std20.Caption, 20); Tab(105); Valor20.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(21); Tab(25); Descri21.Caption; Tab(80); Left$(Std21.Caption, 20); Tab(105); Valor21.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(22); Tab(25); Descri22.Caption; Tab(80); Left$(Std22.Caption, 20); Tab(105); Valor22.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(23); Tab(25); Descri23.Caption; Tab(80); Left$(Std23.Caption, 20); Tab(105); Valor23.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(24); Tab(25); Descri24.Caption; Tab(80); Left$(Std24.Caption, 20); Tab(105); Valor24.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(25); Tab(25); Descri25.Caption; Tab(80); Left$(Std25.Caption, 20); Tab(105); Valor25.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(26); Tab(25); Descri26.Caption; Tab(80); Left$(Std26.Caption, 20); Tab(105); Valor26.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(27); Tab(25); Descri27.Caption; Tab(80); Left$(Std27.Caption, 20); Tab(105); Valor27.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(28); Tab(25); Descri28.Caption; Tab(80); Left$(Std28.Caption, 20); Tab(105); Valor28.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(29); Tab(25); Descri29.Caption; Tab(80); Left$(Std29.Caption, 20); Tab(105); Valor29.Text
    Printer.Print Tab(1); "Ensayo"; Tab(15); ZEnsayo(30); Tab(25); Descri30.Caption; Tab(80); Left$(Std30.Caption, 20); Tab(105); Valor30.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Observaciones"; Tab(20); Ensayo.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Aspecto.Text
    Printer.Print Tab(1); "Observaciones"; Tab(20); Observaciones.Text
    Printer.Print Tab(1); "Confecciono"; Tab(20); Confecciono.Text
    Printer.Print Tab(1); ""
    Printer.Print Tab(1); "Liberada"; Tab(30); Pusing("###,###", Val(Liberada.Text))
    Printer.Print Tab(1); "Devuelta"; Tab(30); Pusing("###,###", Val(Devuelta.Text))
    Printer.Print Tab(1); "Nro Rechazo"; Tab(30); Pusing("######", Val(NroRechazo.Text))
    Printer.Print Tab(1); ""
    
    Printer.EndDoc
    

End Sub

Private Sub ModificaPrueba()

    ZSql = ""
    ZSql = ZSql + "UPDATE PrueArt SET "
    ZSql = ZSql + "Ensayo = " + "'" + Ensayo.Text + "',"
    ZSql = ZSql + "Aspecto = " + "'" + Aspecto.Text + "',"
    ZSql = ZSql + "Observaciones = " + "'" + Observaciones.Text + "',"
    ZSql = ZSql + "Confecciono = " + "'" + Confecciono.Text + "',"
    ZSql = ZSql + "Valor1 = " + "'" + Valor1.Text + "',"
    ZSql = ZSql + "Valor2 = " + "'" + Valor2.Text + "',"
    ZSql = ZSql + "Valor3 = " + "'" + Valor3.Text + "',"
    ZSql = ZSql + "Valor4 = " + "'" + Valor4.Text + "',"
    ZSql = ZSql + "Valor5 = " + "'" + Valor5.Text + "',"
    ZSql = ZSql + "Valor6 = " + "'" + Valor6.Text + "',"
    ZSql = ZSql + "Valor7 = " + "'" + Valor7.Text + "',"
    ZSql = ZSql + "Valor8 = " + "'" + Valor8.Text + "',"
    ZSql = ZSql + "Valor9 = " + "'" + Valor9.Text + "',"
    ZSql = ZSql + "Valor10 = " + "'" + Valor10.Text + "',"
    ZSql = ZSql + "Valor11 = " + "'" + Valor11.Text + "',"
    ZSql = ZSql + "Valor12 = " + "'" + Valor12.Text + "',"
    ZSql = ZSql + "Valor13 = " + "'" + Valor13.Text + "',"
    ZSql = ZSql + "Valor14 = " + "'" + Valor14.Text + "',"
    ZSql = ZSql + "Valor15 = " + "'" + Valor15.Text + "',"
    ZSql = ZSql + "Valor16 = " + "'" + Valor16.Text + "',"
    ZSql = ZSql + "Valor17 = " + "'" + Valor17.Text + "',"
    ZSql = ZSql + "Valor18 = " + "'" + Valor18.Text + "',"
    ZSql = ZSql + "Valor19 = " + "'" + Valor19.Text + "',"
    ZSql = ZSql + "Valor20 = " + "'" + Valor20.Text + "',"
    ZSql = ZSql + "Valor21 = " + "'" + Valor21.Text + "',"
    ZSql = ZSql + "Valor22 = " + "'" + Valor22.Text + "',"
    ZSql = ZSql + "Valor23 = " + "'" + Valor23.Text + "',"
    ZSql = ZSql + "Valor24 = " + "'" + Valor24.Text + "',"
    ZSql = ZSql + "Valor25 = " + "'" + Valor25.Text + "',"
    ZSql = ZSql + "Valor26 = " + "'" + Valor26.Text + "',"
    ZSql = ZSql + "Valor27 = " + "'" + Valor27.Text + "',"
    ZSql = ZSql + "Valor28 = " + "'" + Valor28.Text + "',"
    ZSql = ZSql + "Valor29 = " + "'" + Valor29.Text + "',"
    ZSql = ZSql + "Valor30 = " + "'" + Valor30.Text + "',"
    ZSql = ZSql + "ValorNumero1 = " + "'" + ValorNumero1.Text + "',"
    ZSql = ZSql + "ValorNumero2 = " + "'" + ValorNumero2.Text + "',"
    ZSql = ZSql + "ValorNumero3 = " + "'" + ValorNumero3.Text + "',"
    ZSql = ZSql + "ValorNumero4 = " + "'" + ValorNumero4.Text + "',"
    ZSql = ZSql + "ValorNumero5 = " + "'" + ValorNumero5.Text + "',"
    ZSql = ZSql + "ValorNumero6 = " + "'" + ValorNumero6.Text + "',"
    ZSql = ZSql + "ValorNumero7 = " + "'" + ValorNumero7.Text + "',"
    ZSql = ZSql + "ValorNumero8 = " + "'" + ValorNumero8.Text + "',"
    ZSql = ZSql + "ValorNumero9 = " + "'" + ValorNumero9.Text + "',"
    ZSql = ZSql + "ValorNumero10 = " + "'" + ValorNumero10.Text + "',"
    ZSql = ZSql + "ValorNumero11 = " + "'" + ValorNumero11.Text + "',"
    ZSql = ZSql + "ValorNumero12 = " + "'" + ValorNumero12.Text + "',"
    ZSql = ZSql + "ValorNumero13 = " + "'" + ValorNumero13.Text + "',"
    ZSql = ZSql + "ValorNumero14 = " + "'" + ValorNumero14.Text + "',"
    ZSql = ZSql + "ValorNumero15 = " + "'" + ValorNumero15.Text + "',"
    ZSql = ZSql + "ValorNumero16 = " + "'" + ValorNumero16.Text + "',"
    ZSql = ZSql + "ValorNumero17 = " + "'" + ValorNumero17.Text + "',"
    ZSql = ZSql + "ValorNumero18 = " + "'" + ValorNumero18.Text + "',"
    ZSql = ZSql + "ValorNumero19 = " + "'" + ValorNumero19.Text + "',"
    ZSql = ZSql + "ValorNumero20 = " + "'" + ValorNumero20.Text + "'"
    ZSql = ZSql + "ValorNumero21 = " + "'" + ValorNumero21.Text + "'"
    ZSql = ZSql + "ValorNumero22 = " + "'" + ValorNumero22.Text + "'"
    ZSql = ZSql + "ValorNumero23 = " + "'" + ValorNumero23.Text + "'"
    ZSql = ZSql + "ValorNumero24 = " + "'" + ValorNumero24.Text + "'"
    ZSql = ZSql + "ValorNumero25 = " + "'" + ValorNumero25.Text + "'"
    ZSql = ZSql + "ValorNumero26 = " + "'" + ValorNumero26.Text + "'"
    ZSql = ZSql + "ValorNumero27 = " + "'" + ValorNumero27.Text + "'"
    ZSql = ZSql + "ValorNumero28 = " + "'" + ValorNumero28.Text + "'"
    ZSql = ZSql + "ValorNumero29 = " + "'" + ValorNumero29.Text + "'"
    ZSql = ZSql + "ValorNumero30 = " + "'" + ValorNumero30.Text + "'"
    ZSql = ZSql + " Where Prueba = " + "'" + "1" + Partida.Text + "'"
    
    spPrueart = ZSql
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    
    Call CmdLimpiar_Click
    Producto.SetFocus
End Sub

Private Sub Modifica_Click()
    WProceso = 1
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub Informe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select * "
        ZSql = ZSql + " FROM Informe"
        ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
        ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
        spInforme = ZSql
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            rstInforme.Close
            Orden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        Llave = "N"
        Sql1 = "Select * "
        Sql2 = " FROM Orden"
        Sql3 = " Where Orden = " + "'" + Orden.Text + "'"
        Sql4 = " and Articulo = " + "'" + Producto.Text + "'"
        spOrden = Sql1 + Sql2 + Sql3 + Sql4
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            Llave = "S"
            If rstOrden!Recibida = 0 Then
                Llave = "X"
                    Else
                WOrigen = rstOrden!Origen
            End If
            rstOrden.Close
        End If
    
        Select Case Llave
            Case "S"
                ZVencimiento = "  /  /    "
                ZSql = ""
                ZSql = ZSql + "Select * "
                ZSql = ZSql + " FROM Informe"
                ZSql = ZSql + " Where Informe = " + "'" + Informe.Text + "'"
                ZSql = ZSql + " and Orden = " + "'" + Orden.Text + "'"
                ZSql = ZSql + " and Articulo = " + "'" + Producto.Text + "'"
                spInforme = ZSql
                Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
                If rstInforme.RecordCount > 0 Then
                    Informe = rstInforme!Informe
                    XEnvase = rstInforme!Envase
                    XCertificado1 = IIf(IsNull(rstInforme!Certificado1), "0", rstInforme!Certificado1)
                    XCertificado2 = IIf(IsNull(rstInforme!Certificado2), "", RTrim(rstInforme!Certificado2))
                    XEstado1 = IIf(IsNull(rstInforme!Estado1), "0", rstInforme!Estado1)
                    XEstado2 = IIf(IsNull(rstInforme!Estado2), "", RTrim(rstInforme!Estado2))
                    ZVencimiento = IIf(IsNull(rstInforme!fechavencimiento), "  /  /    ", rstInforme!fechavencimiento)
                    ZNroDespacho = IIf(IsNull(rstInforme!NroDespacho), "", rstInforme!NroDespacho)
                    ZProcedencia = IIf(IsNull(rstInforme!Procedencia), "", rstInforme!Procedencia)
                    rstInforme.Close
                End If
                    
                If XCertificado1 = 0 Then
                    CertificadoNo.Value = 1
                    CertificadoSi.Value = 0
                        Else
                    CertificadoNo.Value = 0
                    CertificadoSi.Value = 1
                End If
                If XEstado1 = 0 Then
                    EstadoNo.Value = 1
                    EstadoSi.Value = 0
                        Else
                    EstadoNo.Value = 0
                    EstadoSi.Value = 1
                End If
                
                Vencimiento.Text = ZVencimiento
                    
                Certificado2.Text = XCertificado2
                Estado2.Text = XEstado2
                IngresaEstado.Height = 3800
                IngresaEstado.Left = 2640
                IngresaEstado.Top = 1500
                IngresaEstado.Width = 6335
                IngresaEstado.Visible = True
                    
            Case "N"
                m$ = "Orden de Compra o articulo inexistente en la orden de compra especificada"
                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                Orden.SetFocus
                
            Case "X"
                m$ = "Orden de compra sin la actualizacion de Informe de Recepcion"
                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                Orden.SetFocus
                
            Case Else
        End Select
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Listado_Click()
    Desde.Text = "  -   -   "
    Desdefec.Text = "  /  /    "
    Hastafec.Text = "  /  /    "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Aprobado.Value = True
    Rechazo.Value = False
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub

Private Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desdefec.SetFocus
    End If
End Sub

Private Sub Desdefec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastafec.SetFocus
    End If
End Sub

Private Sub Hastafec_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.SetFocus
    End If
End Sub

Private Sub Revalida_Click()
    If Val(Partida.Text) <> 0 Then
        ZLoteRevalida = Partida.Text
        ZFechaRevalida = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        ZArticuloRevalida = Producto.Text
        ZDesArticuloRevalida = Descriprod.Caption
        PrgRevalida.Show
    End If
End Sub

Private Sub RevalidaAnterior_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Revalida"
        ZSql = ZSql + " Where Revalida.Revalida = " + "'" + RevalidaAnterior.Text + "'"
        ZSql = ZSql + " and Revalida.Articulo = " + "'" + Producto.Text + "'"
        spRevalida = ZSql
        Set rstRevalida = db.OpenRecordset(spRevalida, dbOpenSnapshot, dbSQLPassThrough)
        If rstRevalida.RecordCount > 0 Then
            rstRevalida.Close
            ZLoteRevalida = Partida.Text
            ZArticuloRevalida = Producto.Text
            ZDesArticuloRevalida = Descriprod.Caption
            ZNroRevalida = RevalidaAnterior.Text
            PrgRevalidaConsulta.Show
        End If
    End If
End Sub

















Private Sub ValorNumero1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(1)), ".")
            ZNumeII = Len(Trim(ZDesde(1)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero1.Text = Pusing("###,###.#", ValorNumero1.Text)
                Case 2
                    ValorNumero1.Text = Pusing("###,###.##", ValorNumero1.Text)
                Case 3
                    ValorNumero1.Text = Pusing("###,###.###", ValorNumero1.Text)
                Case 4
                    ValorNumero1.Text = Pusing("###,###.####", ValorNumero1.Text)
                Case 5
                    ValorNumero1.Text = Pusing("###,###.#####", ValorNumero1.Text)
                Case 6
                    ValorNumero1.Text = Pusing("###,###.######", ValorNumero1.Text)
                Case Else
                    ValorNumero1.Text = Pusing("###,###", ValorNumero1.Text)
            End Select
            
            Valor1.Text = ValorNumero1.Text + " " + ZUnidad(1)
            
            ValorNumero2.SetFocus
            
                Else
                
            If ValorNumero1.Text = "S" Or ValorNumero1.Text = "N" Then
                If ValorNumero1.Text = "S" Then
                    Valor1.Text = "Cumple"
                        Else
                    Valor1.Text = "No Cumple"
                End If
                ValorNumero2.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero1.Text = ""
    End If
    
    If Val(ZDesde(1)) <> 0 Or Val(ZHasta(1)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(2)), ".")
            ZNumeII = Len(Trim(ZDesde(2)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero2.Text = Pusing("###,###.#", ValorNumero2.Text)
                Case 2
                    ValorNumero2.Text = Pusing("###,###.##", ValorNumero2.Text)
                Case 3
                    ValorNumero2.Text = Pusing("###,###.###", ValorNumero2.Text)
                Case 4
                    ValorNumero2.Text = Pusing("###,###.####", ValorNumero2.Text)
                Case 5
                    ValorNumero2.Text = Pusing("###,###.#####", ValorNumero2.Text)
                Case 6
                    ValorNumero2.Text = Pusing("###,###.######", ValorNumero2.Text)
                Case Else
                    ValorNumero2.Text = Pusing("###,###", ValorNumero2.Text)
            End Select
            
            Valor2.Text = ValorNumero2.Text + " " + ZUnidad(2)
            
            ValorNumero3.SetFocus
            
                Else
                
            If ValorNumero2.Text = "S" Or ValorNumero2.Text = "N" Then
                If ValorNumero2.Text = "S" Then
                    Valor2.Text = "Cumple"
                        Else
                    Valor2.Text = "No Cumple"
                End If
                ValorNumero3.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero2.Text = ""
    End If
    
    If Val(ZDesde(2)) <> 0 Or Val(ZHasta(2)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(3)), ".")
            ZNumeII = Len(Trim(ZDesde(3)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero3.Text = Pusing("###,###.#", ValorNumero3.Text)
                Case 2
                    ValorNumero3.Text = Pusing("###,###.##", ValorNumero3.Text)
                Case 3
                    ValorNumero3.Text = Pusing("###,###.###", ValorNumero3.Text)
                Case 4
                    ValorNumero3.Text = Pusing("###,###.####", ValorNumero3.Text)
                Case 5
                    ValorNumero3.Text = Pusing("###,###.#####", ValorNumero3.Text)
                Case 6
                    ValorNumero3.Text = Pusing("###,###.######", ValorNumero3.Text)
                Case Else
                    ValorNumero3.Text = Pusing("###,###", ValorNumero3.Text)
            End Select
            
            Valor3.Text = ValorNumero3.Text + " " + ZUnidad(3)
            
            ValorNumero4.SetFocus
            
                Else
                
            If ValorNumero3.Text = "S" Or ValorNumero3.Text = "N" Then
                If ValorNumero3.Text = "S" Then
                    Valor3.Text = "Cumple"
                        Else
                    Valor3.Text = "No Cumple"
                End If
                ValorNumero4.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero3.Text = ""
    End If
    
    If Val(ZDesde(3)) <> 0 Or Val(ZHasta(3)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" And KeyAscii <> 8 Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(4)), ".")
            ZNumeII = Len(Trim(ZDesde(4)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero4.Text = Pusing("###,###.#", ValorNumero4.Text)
                Case 2
                    ValorNumero4.Text = Pusing("###,###.##", ValorNumero4.Text)
                Case 3
                    ValorNumero4.Text = Pusing("###,###.###", ValorNumero4.Text)
                Case 4
                    ValorNumero4.Text = Pusing("###,###.####", ValorNumero4.Text)
                Case 5
                    ValorNumero4.Text = Pusing("###,###.#####", ValorNumero4.Text)
                Case 6
                    ValorNumero4.Text = Pusing("###,###.######", ValorNumero4.Text)
                Case Else
                    ValorNumero4.Text = Pusing("###,###", ValorNumero4.Text)
            End Select
            
            Valor4.Text = ValorNumero4.Text + " " + ZUnidad(4)
            
            ValorNumero5.SetFocus
            
                Else
                
            If ValorNumero4.Text = "S" Or ValorNumero4.Text = "N" Then
                If ValorNumero4.Text = "S" Then
                    Valor4.Text = "Cumple"
                        Else
                    Valor4.Text = "No Cumple"
                End If
                ValorNumero5.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero4.Text = ""
    End If
    
    If Val(ZDesde(4)) <> 0 Or Val(ZHasta(4)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub




Private Sub ValorNumero5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(5)), ".")
            ZNumeII = Len(Trim(ZDesde(5)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero5.Text = Pusing("###,###.#", ValorNumero5.Text)
                Case 2
                    ValorNumero5.Text = Pusing("###,###.##", ValorNumero5.Text)
                Case 3
                    ValorNumero5.Text = Pusing("###,###.###", ValorNumero5.Text)
                Case 4
                    ValorNumero5.Text = Pusing("###,###.####", ValorNumero5.Text)
                Case 5
                    ValorNumero5.Text = Pusing("###,###.#####", ValorNumero5.Text)
                Case 6
                    ValorNumero5.Text = Pusing("###,###.######", ValorNumero5.Text)
                Case Else
                    ValorNumero5.Text = Pusing("###,###", ValorNumero5.Text)
            End Select
            
            Valor5.Text = ValorNumero5.Text + " " + ZUnidad(5)
            
            ValorNumero6.SetFocus
            
                Else
                
            If ValorNumero5.Text = "S" Or ValorNumero5.Text = "N" Then
                If ValorNumero5.Text = "S" Then
                    Valor5.Text = "Cumple"
                        Else
                    Valor5.Text = "No Cumple"
                End If
                ValorNumero6.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero5.Text = ""
    End If
    
    If Val(ZDesde(5)) <> 0 Or Val(ZHasta(5)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(6)), ".")
            ZNumeII = Len(Trim(ZDesde(6)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero6.Text = Pusing("###,###.#", ValorNumero6.Text)
                Case 2
                    ValorNumero6.Text = Pusing("###,###.##", ValorNumero6.Text)
                Case 3
                    ValorNumero6.Text = Pusing("###,###.###", ValorNumero6.Text)
                Case 4
                    ValorNumero6.Text = Pusing("###,###.####", ValorNumero6.Text)
                Case 5
                    ValorNumero6.Text = Pusing("###,###.#####", ValorNumero6.Text)
                Case 6
                    ValorNumero6.Text = Pusing("###,###.######", ValorNumero6.Text)
                Case Else
                    ValorNumero6.Text = Pusing("###,###", ValorNumero6.Text)
            End Select
            
            Valor6.Text = ValorNumero6.Text + " " + ZUnidad(6)
            
            ValorNumero7.SetFocus
            
                Else
                
            If ValorNumero6.Text = "S" Or ValorNumero6.Text = "N" Then
                If ValorNumero6.Text = "S" Then
                    Valor6.Text = "Cumple"
                        Else
                    Valor6.Text = "No Cumple"
                End If
                ValorNumero7.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero6.Text = ""
    End If
    
    If Val(ZDesde(6)) <> 0 Or Val(ZHasta(6)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(7)), ".")
            ZNumeII = Len(Trim(ZDesde(7)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero7.Text = Pusing("###,###.#", ValorNumero7.Text)
                Case 2
                    ValorNumero7.Text = Pusing("###,###.##", ValorNumero7.Text)
                Case 3
                    ValorNumero7.Text = Pusing("###,###.###", ValorNumero7.Text)
                Case 4
                    ValorNumero7.Text = Pusing("###,###.####", ValorNumero7.Text)
                Case 5
                    ValorNumero7.Text = Pusing("###,###.#####", ValorNumero7.Text)
                Case 6
                    ValorNumero7.Text = Pusing("###,###.######", ValorNumero7.Text)
                Case Else
                    ValorNumero7.Text = Pusing("###,###", ValorNumero7.Text)
            End Select
            
            Valor7.Text = ValorNumero7.Text + " " + ZUnidad(7)
            
            ValorNumero8.SetFocus
            
                Else
                
            If ValorNumero7.Text = "S" Or ValorNumero7.Text = "N" Then
                If ValorNumero7.Text = "S" Then
                    Valor7.Text = "Cumple"
                        Else
                    Valor7.Text = "No Cumple"
                End If
                ValorNumero8.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero7.Text = ""
    End If
    
    If Val(ZDesde(7)) <> 0 Or Val(ZHasta(7)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(8)), ".")
            ZNumeII = Len(Trim(ZDesde(8)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero8.Text = Pusing("###,###.#", ValorNumero8.Text)
                Case 2
                    ValorNumero8.Text = Pusing("###,###.##", ValorNumero8.Text)
                Case 3
                    ValorNumero8.Text = Pusing("###,###.###", ValorNumero8.Text)
                Case 4
                    ValorNumero8.Text = Pusing("###,###.####", ValorNumero8.Text)
                Case 5
                    ValorNumero8.Text = Pusing("###,###.#####", ValorNumero8.Text)
                Case 6
                    ValorNumero8.Text = Pusing("###,###.######", ValorNumero8.Text)
                Case Else
                    ValorNumero8.Text = Pusing("###,###", ValorNumero8.Text)
            End Select
            
            Valor8.Text = ValorNumero8.Text + " " + ZUnidad(8)
            
            ValorNumero9.SetFocus
            
                Else
                
            If ValorNumero8.Text = "S" Or ValorNumero8.Text = "N" Then
                If ValorNumero8.Text = "S" Then
                    Valor8.Text = "Cumple"
                        Else
                    Valor8.Text = "No Cumple"
                End If
                ValorNumero9.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero8.Text = ""
    End If
    
    If Val(ZDesde(8)) <> 0 Or Val(ZHasta(8)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(9)), ".")
            ZNumeII = Len(Trim(ZDesde(9)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero9.Text = Pusing("###,###.#", ValorNumero9.Text)
                Case 2
                    ValorNumero9.Text = Pusing("###,###.##", ValorNumero9.Text)
                Case 3
                    ValorNumero9.Text = Pusing("###,###.###", ValorNumero9.Text)
                Case 4
                    ValorNumero9.Text = Pusing("###,###.####", ValorNumero9.Text)
                Case 5
                    ValorNumero9.Text = Pusing("###,###.#####", ValorNumero9.Text)
                Case 6
                    ValorNumero9.Text = Pusing("###,###.######", ValorNumero9.Text)
                Case Else
                    ValorNumero9.Text = Pusing("###,###", ValorNumero9.Text)
            End Select
            
            Valor9.Text = ValorNumero9.Text + " " + ZUnidad(9)
            
            ValorNumero10.SetFocus
            
                Else
                
            If ValorNumero9.Text = "S" Or ValorNumero9.Text = "N" Then
                If ValorNumero9.Text = "S" Then
                    Valor9.Text = "Cumple"
                        Else
                    Valor9.Text = "No Cumple"
                End If
                ValorNumero10.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero9.Text = ""
    End If
    
    If Val(ZDesde(9)) <> 0 Or Val(ZHasta(9)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(10)), ".")
            ZNumeII = Len(Trim(ZDesde(10)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero10.Text = Pusing("###,###.#", ValorNumero10.Text)
                Case 2
                    ValorNumero10.Text = Pusing("###,###.##", ValorNumero10.Text)
                Case 3
                    ValorNumero10.Text = Pusing("###,###.###", ValorNumero10.Text)
                Case 4
                    ValorNumero10.Text = Pusing("###,###.####", ValorNumero10.Text)
                Case 5
                    ValorNumero10.Text = Pusing("###,###.#####", ValorNumero10.Text)
                Case 6
                    ValorNumero10.Text = Pusing("###,###.######", ValorNumero10.Text)
                Case Else
                    ValorNumero10.Text = Pusing("###,###", ValorNumero10.Text)
            End Select
            
            Valor10.Text = ValorNumero10.Text + " " + ZUnidad(10)
            
            SSTab1.Tab = 1
            ValorNumero11.SetFocus
            
                Else
                
            If ValorNumero10.Text = "S" Or ValorNumero10.Text = "N" Then
                If ValorNumero10.Text = "S" Then
                    Valor10.Text = "Cumple"
                        Else
                    Valor10.Text = "No Cumple"
                End If
                SSTab1.Tab = 1
                ValorNumero11.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero10.Text = ""
    End If
    
    If Val(ZDesde(10)) <> 0 Or Val(ZHasta(10)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(11)) <> 0 Or Val(ZHasta(11)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(11)), ".")
            ZNumeII = Len(Trim(ZDesde(11)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero11.Text = Pusing("###,###.#", ValorNumero11.Text)
                Case 2
                    ValorNumero11.Text = Pusing("###,###.##", ValorNumero11.Text)
                Case 3
                    ValorNumero11.Text = Pusing("###,###.###", ValorNumero11.Text)
                Case 4
                    ValorNumero11.Text = Pusing("###,###.####", ValorNumero11.Text)
                Case 5
                    ValorNumero11.Text = Pusing("###,###.#####", ValorNumero11.Text)
                Case 6
                    ValorNumero11.Text = Pusing("###,###.######", ValorNumero11.Text)
                Case Else
                    ValorNumero11.Text = Pusing("###,###", ValorNumero11.Text)
            End Select
            
            Valor11.Text = ValorNumero11.Text + " " + ZUnidad(11)
            
            ValorNumero12.SetFocus
            
                Else
                
            If ValorNumero11.Text = "S" Or ValorNumero11.Text = "N" Then
                If ValorNumero11.Text = "S" Then
                    Valor11.Text = "Cumple"
                        Else
                    Valor11.Text = "No Cumple"
                End If
                ValorNumero12.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero11.Text = ""
    End If
    
    If Val(ZDesde(11)) <> 0 Or Val(ZHasta(11)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(12)) <> 0 Or Val(ZHasta(12)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(12)), ".")
            ZNumeII = Len(Trim(ZDesde(12)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero12.Text = Pusing("###,###.#", ValorNumero12.Text)
                Case 2
                    ValorNumero12.Text = Pusing("###,###.##", ValorNumero12.Text)
                Case 3
                    ValorNumero12.Text = Pusing("###,###.###", ValorNumero12.Text)
                Case 4
                    ValorNumero12.Text = Pusing("###,###.####", ValorNumero12.Text)
                Case 5
                    ValorNumero12.Text = Pusing("###,###.#####", ValorNumero12.Text)
                Case 6
                    ValorNumero12.Text = Pusing("###,###.######", ValorNumero12.Text)
                Case Else
                    ValorNumero12.Text = Pusing("###,###", ValorNumero12.Text)
            End Select
            
            Valor12.Text = ValorNumero12.Text + " " + ZUnidad(12)
            
            ValorNumero13.SetFocus
            
                Else
                
            If ValorNumero12.Text = "S" Or ValorNumero12.Text = "N" Then
                If ValorNumero12.Text = "S" Then
                    Valor12.Text = "Cumple"
                        Else
                    Valor12.Text = "No Cumple"
                End If
                ValorNumero13.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero12.Text = ""
    End If
    
    If Val(ZDesde(12)) <> 0 Or Val(ZHasta(12)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(13)) <> 0 Or Val(ZHasta(13)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(13)), ".")
            ZNumeII = Len(Trim(ZDesde(13)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero13.Text = Pusing("###,###.#", ValorNumero13.Text)
                Case 2
                    ValorNumero13.Text = Pusing("###,###.##", ValorNumero13.Text)
                Case 3
                    ValorNumero13.Text = Pusing("###,###.###", ValorNumero13.Text)
                Case 4
                    ValorNumero13.Text = Pusing("###,###.####", ValorNumero13.Text)
                Case 5
                    ValorNumero13.Text = Pusing("###,###.#####", ValorNumero13.Text)
                Case 6
                    ValorNumero13.Text = Pusing("###,###.######", ValorNumero13.Text)
                Case Else
                    ValorNumero13.Text = Pusing("###,###", ValorNumero13.Text)
            End Select
            
            Valor13.Text = ValorNumero13.Text + " " + ZUnidad(13)
            
            ValorNumero14.SetFocus
            
                Else
                
            If ValorNumero13.Text = "S" Or ValorNumero13.Text = "N" Then
                If ValorNumero13.Text = "S" Then
                    Valor13.Text = "Cumple"
                        Else
                    Valor13.Text = "No Cumple"
                End If
                ValorNumero14.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero13.Text = ""
    End If
    
    If Val(ZDesde(13)) <> 0 Or Val(ZHasta(13)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(14)) <> 0 Or Val(ZHasta(14)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(14)), ".")
            ZNumeII = Len(Trim(ZDesde(14)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero14.Text = Pusing("###,###.#", ValorNumero14.Text)
                Case 2
                    ValorNumero14.Text = Pusing("###,###.##", ValorNumero14.Text)
                Case 3
                    ValorNumero14.Text = Pusing("###,###.###", ValorNumero14.Text)
                Case 4
                    ValorNumero14.Text = Pusing("###,###.####", ValorNumero14.Text)
                Case 5
                    ValorNumero14.Text = Pusing("###,###.#####", ValorNumero14.Text)
                Case 6
                    ValorNumero14.Text = Pusing("###,###.######", ValorNumero14.Text)
                Case Else
                    ValorNumero14.Text = Pusing("###,###", ValorNumero14.Text)
            End Select
            
            Valor14.Text = ValorNumero14.Text + " " + ZUnidad(14)
            
            ValorNumero15.SetFocus
            
                Else
                
            If ValorNumero14.Text = "S" Or ValorNumero14.Text = "N" Then
                If ValorNumero14.Text = "S" Then
                    Valor14.Text = "Cumple"
                        Else
                    Valor14.Text = "No Cumple"
                End If
                ValorNumero15.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero14.Text = ""
    End If
    
    If Val(ZDesde(14)) <> 0 Or Val(ZHasta(14)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(15)) <> 0 Or Val(ZHasta(15)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(15)), ".")
            ZNumeII = Len(Trim(ZDesde(15)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero15.Text = Pusing("###,###.#", ValorNumero15.Text)
                Case 2
                    ValorNumero15.Text = Pusing("###,###.##", ValorNumero15.Text)
                Case 3
                    ValorNumero15.Text = Pusing("###,###.###", ValorNumero15.Text)
                Case 4
                    ValorNumero15.Text = Pusing("###,###.####", ValorNumero15.Text)
                Case 5
                    ValorNumero15.Text = Pusing("###,###.#####", ValorNumero15.Text)
                Case 6
                    ValorNumero15.Text = Pusing("###,###.######", ValorNumero15.Text)
                Case Else
                    ValorNumero15.Text = Pusing("###,###", ValorNumero15.Text)
            End Select
            
            Valor15.Text = ValorNumero15.Text + " " + ZUnidad(15)
            
            ValorNumero16.SetFocus
            
                Else
                
            If ValorNumero15.Text = "S" Or ValorNumero15.Text = "N" Then
                If ValorNumero15.Text = "S" Then
                    Valor15.Text = "Cumple"
                        Else
                    Valor15.Text = "No Cumple"
                End If
                ValorNumero16.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero15.Text = ""
    End If
    
    If Val(ZDesde(15)) <> 0 Or Val(ZHasta(15)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(16)) <> 0 Or Val(ZHasta(16)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(16)), ".")
            ZNumeII = Len(Trim(ZDesde(16)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero16.Text = Pusing("###,###.#", ValorNumero16.Text)
                Case 2
                    ValorNumero16.Text = Pusing("###,###.##", ValorNumero16.Text)
                Case 3
                    ValorNumero16.Text = Pusing("###,###.###", ValorNumero16.Text)
                Case 4
                    ValorNumero16.Text = Pusing("###,###.####", ValorNumero16.Text)
                Case 5
                    ValorNumero16.Text = Pusing("###,###.#####", ValorNumero16.Text)
                Case 6
                    ValorNumero16.Text = Pusing("###,###.######", ValorNumero16.Text)
                Case Else
                    ValorNumero16.Text = Pusing("###,###", ValorNumero16.Text)
            End Select
            
            Valor16.Text = ValorNumero16.Text + " " + ZUnidad(16)
            
            ValorNumero17.SetFocus
            
                Else
                
            If ValorNumero16.Text = "S" Or ValorNumero16.Text = "N" Then
                If ValorNumero16.Text = "S" Then
                    Valor16.Text = "Cumple"
                        Else
                    Valor16.Text = "No Cumple"
                End If
                ValorNumero17.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero16.Text = ""
    End If
    
    If Val(ZDesde(16)) <> 0 Or Val(ZHasta(16)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(17)) <> 0 Or Val(ZHasta(17)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(17)), ".")
            ZNumeII = Len(Trim(ZDesde(17)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero17.Text = Pusing("###,###.#", ValorNumero17.Text)
                Case 2
                    ValorNumero17.Text = Pusing("###,###.##", ValorNumero17.Text)
                Case 3
                    ValorNumero17.Text = Pusing("###,###.###", ValorNumero17.Text)
                Case 4
                    ValorNumero17.Text = Pusing("###,###.####", ValorNumero17.Text)
                Case 5
                    ValorNumero17.Text = Pusing("###,###.#####", ValorNumero17.Text)
                Case 6
                    ValorNumero17.Text = Pusing("###,###.######", ValorNumero17.Text)
                Case Else
                    ValorNumero17.Text = Pusing("###,###", ValorNumero17.Text)
            End Select
            
            Valor17.Text = ValorNumero17.Text + " " + ZUnidad(17)
            
            ValorNumero18.SetFocus
            
                Else
                
            If ValorNumero17.Text = "S" Or ValorNumero17.Text = "N" Then
                If ValorNumero17.Text = "S" Then
                    Valor17.Text = "Cumple"
                        Else
                    Valor17.Text = "No Cumple"
                End If
                ValorNumero18.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero17.Text = ""
    End If
    
    If Val(ZDesde(17)) <> 0 Or Val(ZHasta(17)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(18)) <> 0 Or Val(ZHasta(18)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(18)), ".")
            ZNumeII = Len(Trim(ZDesde(18)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero18.Text = Pusing("###,###.#", ValorNumero18.Text)
                Case 2
                    ValorNumero18.Text = Pusing("###,###.##", ValorNumero18.Text)
                Case 3
                    ValorNumero18.Text = Pusing("###,###.###", ValorNumero18.Text)
                Case 4
                    ValorNumero18.Text = Pusing("###,###.####", ValorNumero18.Text)
                Case 5
                    ValorNumero18.Text = Pusing("###,###.#####", ValorNumero18.Text)
                Case 6
                    ValorNumero18.Text = Pusing("###,###.######", ValorNumero18.Text)
                Case Else
                    ValorNumero18.Text = Pusing("###,###", ValorNumero18.Text)
            End Select
            
            Valor18.Text = ValorNumero18.Text + " " + ZUnidad(18)
            
            ValorNumero19.SetFocus
            
                Else
                
            If ValorNumero18.Text = "S" Or ValorNumero18.Text = "N" Then
                If ValorNumero18.Text = "S" Then
                    Valor18.Text = "Cumple"
                        Else
                    Valor18.Text = "No Cumple"
                End If
                ValorNumero19.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero18.Text = ""
    End If
    
    If Val(ZDesde(18)) <> 0 Or Val(ZHasta(18)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(19)) <> 0 Or Val(ZHasta(19)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(19)), ".")
            ZNumeII = Len(Trim(ZDesde(19)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero19.Text = Pusing("###,###.#", ValorNumero19.Text)
                Case 2
                    ValorNumero19.Text = Pusing("###,###.##", ValorNumero19.Text)
                Case 3
                    ValorNumero19.Text = Pusing("###,###.###", ValorNumero19.Text)
                Case 4
                    ValorNumero19.Text = Pusing("###,###.####", ValorNumero19.Text)
                Case 5
                    ValorNumero19.Text = Pusing("###,###.#####", ValorNumero19.Text)
                Case 6
                    ValorNumero19.Text = Pusing("###,###.######", ValorNumero19.Text)
                Case Else
                    ValorNumero19.Text = Pusing("###,###", ValorNumero19.Text)
            End Select
            
            Valor19.Text = ValorNumero19.Text + " " + ZUnidad(19)
            
            ValorNumero20.SetFocus
            
                Else
                
            If ValorNumero19.Text = "S" Or ValorNumero19.Text = "N" Then
                If ValorNumero19.Text = "S" Then
                    Valor19.Text = "Cumple"
                        Else
                    Valor19.Text = "No Cumple"
                End If
                ValorNumero20.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero19.Text = ""
    End If
    
    If Val(ZDesde(19)) <> 0 Or Val(ZHasta(19)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub


Private Sub ValorNumero20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(20)) <> 0 Or Val(ZHasta(20)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(20)), ".")
            ZNumeII = Len(Trim(ZDesde(20)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero20.Text = Pusing("###,###.#", ValorNumero20.Text)
                Case 2
                    ValorNumero20.Text = Pusing("###,###.##", ValorNumero20.Text)
                Case 3
                    ValorNumero20.Text = Pusing("###,###.###", ValorNumero20.Text)
                Case 4
                    ValorNumero20.Text = Pusing("###,###.####", ValorNumero20.Text)
                Case 5
                    ValorNumero20.Text = Pusing("###,###.#####", ValorNumero20.Text)
                Case 6
                    ValorNumero20.Text = Pusing("###,###.######", ValorNumero20.Text)
                Case Else
                    ValorNumero20.Text = Pusing("###,###", ValorNumero20.Text)
            End Select
            
            Valor20.Text = ValorNumero20.Text + " " + ZUnidad(20)
            
            SSTab1.Tab = 2
            ValorNumero21.SetFocus
            
                Else
                
            If ValorNumero20.Text = "S" Or ValorNumero20.Text = "N" Then
                If ValorNumero20.Text = "S" Then
                    Valor20.Text = "Cumple"
                        Else
                    Valor20.Text = "No Cumple"
                End If
                SSTab1.Tab = 2
                ValorNumero21.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero20.Text = ""
    End If
    
    If Val(ZDesde(20)) <> 0 Or Val(ZHasta(20)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub





Private Sub ValorNumero21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(21)) <> 0 Or Val(ZHasta(21)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(21)), ".")
            ZNumeII = Len(Trim(ZDesde(21)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero21.Text = Pusing("###,###.#", ValorNumero21.Text)
                Case 2
                    ValorNumero21.Text = Pusing("###,###.##", ValorNumero21.Text)
                Case 3
                    ValorNumero21.Text = Pusing("###,###.###", ValorNumero21.Text)
                Case 4
                    ValorNumero21.Text = Pusing("###,###.####", ValorNumero21.Text)
                Case 5
                    ValorNumero21.Text = Pusing("###,###.#####", ValorNumero21.Text)
                Case 6
                    ValorNumero21.Text = Pusing("###,###.######", ValorNumero21.Text)
                Case Else
                    ValorNumero21.Text = Pusing("###,###", ValorNumero21.Text)
            End Select
            
            Valor21.Text = ValorNumero21.Text + " " + ZUnidad(21)
            
            ValorNumero22.SetFocus
            
                Else
                
            If ValorNumero21.Text = "S" Or ValorNumero21.Text = "N" Then
                If ValorNumero21.Text = "S" Then
                    Valor21.Text = "Cumple"
                        Else
                    Valor21.Text = "No Cumple"
                End If
                ValorNumero22.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero21.Text = ""
    End If
    
    If Val(ZDesde(21)) <> 0 Or Val(ZHasta(21)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(22)) <> 0 Or Val(ZHasta(22)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(22)), ".")
            ZNumeII = Len(Trim(ZDesde(22)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero22.Text = Pusing("###,###.#", ValorNumero22.Text)
                Case 2
                    ValorNumero22.Text = Pusing("###,###.##", ValorNumero22.Text)
                Case 3
                    ValorNumero22.Text = Pusing("###,###.###", ValorNumero22.Text)
                Case 4
                    ValorNumero22.Text = Pusing("###,###.####", ValorNumero22.Text)
                Case 5
                    ValorNumero22.Text = Pusing("###,###.#####", ValorNumero22.Text)
                Case 6
                    ValorNumero22.Text = Pusing("###,###.######", ValorNumero22.Text)
                Case Else
                    ValorNumero22.Text = Pusing("###,###", ValorNumero22.Text)
            End Select
            
            Valor22.Text = ValorNumero22.Text + " " + ZUnidad(22)
            
            ValorNumero23.SetFocus
            
                Else
                
            If ValorNumero22.Text = "S" Or ValorNumero22.Text = "N" Then
                If ValorNumero22.Text = "S" Then
                    Valor22.Text = "Cumple"
                        Else
                    Valor22.Text = "No Cumple"
                End If
                ValorNumero23.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero22.Text = ""
    End If
    
    If Val(ZDesde(22)) <> 0 Or Val(ZHasta(22)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(23)) <> 0 Or Val(ZHasta(23)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(23)), ".")
            ZNumeII = Len(Trim(ZDesde(23)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero23.Text = Pusing("###,###.#", ValorNumero23.Text)
                Case 2
                    ValorNumero23.Text = Pusing("###,###.##", ValorNumero23.Text)
                Case 3
                    ValorNumero23.Text = Pusing("###,###.###", ValorNumero23.Text)
                Case 4
                    ValorNumero23.Text = Pusing("###,###.####", ValorNumero23.Text)
                Case 5
                    ValorNumero23.Text = Pusing("###,###.#####", ValorNumero23.Text)
                Case 6
                    ValorNumero23.Text = Pusing("###,###.######", ValorNumero23.Text)
                Case Else
                    ValorNumero23.Text = Pusing("###,###", ValorNumero23.Text)
            End Select
            
            Valor23.Text = ValorNumero23.Text + " " + ZUnidad(23)
            
            ValorNumero24.SetFocus
            
                Else
                
            If ValorNumero23.Text = "S" Or ValorNumero23.Text = "N" Then
                If ValorNumero23.Text = "S" Then
                    Valor23.Text = "Cumple"
                        Else
                    Valor23.Text = "No Cumple"
                End If
                ValorNumero24.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero23.Text = ""
    End If
    
    If Val(ZDesde(23)) <> 0 Or Val(ZHasta(23)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(24)) <> 0 Or Val(ZHasta(24)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(24)), ".")
            ZNumeII = Len(Trim(ZDesde(24)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero24.Text = Pusing("###,###.#", ValorNumero24.Text)
                Case 2
                    ValorNumero24.Text = Pusing("###,###.##", ValorNumero24.Text)
                Case 3
                    ValorNumero24.Text = Pusing("###,###.###", ValorNumero24.Text)
                Case 4
                    ValorNumero24.Text = Pusing("###,###.####", ValorNumero24.Text)
                Case 5
                    ValorNumero24.Text = Pusing("###,###.#####", ValorNumero24.Text)
                Case 6
                    ValorNumero24.Text = Pusing("###,###.######", ValorNumero24.Text)
                Case Else
                    ValorNumero24.Text = Pusing("###,###", ValorNumero24.Text)
            End Select
            
            Valor24.Text = ValorNumero24.Text + " " + ZUnidad(24)
            
            ValorNumero25.SetFocus
            
                Else
                
            If ValorNumero24.Text = "S" Or ValorNumero24.Text = "N" Then
                If ValorNumero24.Text = "S" Then
                    Valor24.Text = "Cumple"
                        Else
                    Valor24.Text = "No Cumple"
                End If
                ValorNumero25.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero24.Text = ""
    End If
    
    If Val(ZDesde(24)) <> 0 Or Val(ZHasta(24)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(25)) <> 0 Or Val(ZHasta(25)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(25)), ".")
            ZNumeII = Len(Trim(ZDesde(25)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero25.Text = Pusing("###,###.#", ValorNumero25.Text)
                Case 2
                    ValorNumero25.Text = Pusing("###,###.##", ValorNumero25.Text)
                Case 3
                    ValorNumero25.Text = Pusing("###,###.###", ValorNumero25.Text)
                Case 4
                    ValorNumero25.Text = Pusing("###,###.####", ValorNumero25.Text)
                Case 5
                    ValorNumero25.Text = Pusing("###,###.#####", ValorNumero25.Text)
                Case 6
                    ValorNumero25.Text = Pusing("###,###.######", ValorNumero25.Text)
                Case Else
                    ValorNumero25.Text = Pusing("###,###", ValorNumero25.Text)
            End Select
            
            Valor25.Text = ValorNumero25.Text + " " + ZUnidad(25)
            
            ValorNumero26.SetFocus
            
                Else
                
            If ValorNumero25.Text = "S" Or ValorNumero25.Text = "N" Then
                If ValorNumero25.Text = "S" Then
                    Valor25.Text = "Cumple"
                        Else
                    Valor25.Text = "No Cumple"
                End If
                ValorNumero26.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero25.Text = ""
    End If
    
    If Val(ZDesde(25)) <> 0 Or Val(ZHasta(25)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(26)) <> 0 Or Val(ZHasta(26)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(26)), ".")
            ZNumeII = Len(Trim(ZDesde(26)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero26.Text = Pusing("###,###.#", ValorNumero26.Text)
                Case 2
                    ValorNumero26.Text = Pusing("###,###.##", ValorNumero26.Text)
                Case 3
                    ValorNumero26.Text = Pusing("###,###.###", ValorNumero26.Text)
                Case 4
                    ValorNumero26.Text = Pusing("###,###.####", ValorNumero26.Text)
                Case 5
                    ValorNumero26.Text = Pusing("###,###.#####", ValorNumero26.Text)
                Case 6
                    ValorNumero26.Text = Pusing("###,###.######", ValorNumero26.Text)
                Case Else
                    ValorNumero26.Text = Pusing("###,###", ValorNumero26.Text)
            End Select
            
            Valor26.Text = ValorNumero26.Text + " " + ZUnidad(26)
            
            ValorNumero27.SetFocus
            
                Else
                
            If ValorNumero26.Text = "S" Or ValorNumero26.Text = "N" Then
                If ValorNumero26.Text = "S" Then
                    Valor26.Text = "Cumple"
                        Else
                    Valor26.Text = "No Cumple"
                End If
                ValorNumero27.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero26.Text = ""
    End If
    
    If Val(ZDesde(26)) <> 0 Or Val(ZHasta(26)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(27)) <> 0 Or Val(ZHasta(27)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(27)), ".")
            ZNumeII = Len(Trim(ZDesde(27)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero27.Text = Pusing("###,###.#", ValorNumero27.Text)
                Case 2
                    ValorNumero27.Text = Pusing("###,###.##", ValorNumero27.Text)
                Case 3
                    ValorNumero27.Text = Pusing("###,###.###", ValorNumero27.Text)
                Case 4
                    ValorNumero27.Text = Pusing("###,###.####", ValorNumero27.Text)
                Case 5
                    ValorNumero27.Text = Pusing("###,###.#####", ValorNumero27.Text)
                Case 6
                    ValorNumero27.Text = Pusing("###,###.######", ValorNumero27.Text)
                Case Else
                    ValorNumero27.Text = Pusing("###,###", ValorNumero27.Text)
            End Select
            
            Valor27.Text = ValorNumero27.Text + " " + ZUnidad(27)
            
            ValorNumero28.SetFocus
            
                Else
                
            If ValorNumero27.Text = "S" Or ValorNumero27.Text = "N" Then
                If ValorNumero27.Text = "S" Then
                    Valor27.Text = "Cumple"
                        Else
                    Valor27.Text = "No Cumple"
                End If
                ValorNumero28.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero27.Text = ""
    End If
    
    If Val(ZDesde(27)) <> 0 Or Val(ZHasta(27)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(28)) <> 0 Or Val(ZHasta(28)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(28)), ".")
            ZNumeII = Len(Trim(ZDesde(28)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero28.Text = Pusing("###,###.#", ValorNumero28.Text)
                Case 2
                    ValorNumero28.Text = Pusing("###,###.##", ValorNumero28.Text)
                Case 3
                    ValorNumero28.Text = Pusing("###,###.###", ValorNumero28.Text)
                Case 4
                    ValorNumero28.Text = Pusing("###,###.####", ValorNumero28.Text)
                Case 5
                    ValorNumero28.Text = Pusing("###,###.#####", ValorNumero28.Text)
                Case 6
                    ValorNumero28.Text = Pusing("###,###.######", ValorNumero28.Text)
                Case Else
                    ValorNumero28.Text = Pusing("###,###", ValorNumero28.Text)
            End Select
            
            Valor28.Text = ValorNumero28.Text + " " + ZUnidad(28)
            
            ValorNumero29.SetFocus
            
                Else
                
            If ValorNumero28.Text = "S" Or ValorNumero28.Text = "N" Then
                If ValorNumero28.Text = "S" Then
                    Valor28.Text = "Cumple"
                        Else
                    Valor28.Text = "No Cumple"
                End If
                ValorNumero29.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero28.Text = ""
    End If
    
    If Val(ZDesde(28)) <> 0 Or Val(ZHasta(28)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(29)) <> 0 Or Val(ZHasta(29)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(29)), ".")
            ZNumeII = Len(Trim(ZDesde(29)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero29.Text = Pusing("###,###.#", ValorNumero29.Text)
                Case 2
                    ValorNumero29.Text = Pusing("###,###.##", ValorNumero29.Text)
                Case 3
                    ValorNumero29.Text = Pusing("###,###.###", ValorNumero29.Text)
                Case 4
                    ValorNumero29.Text = Pusing("###,###.####", ValorNumero29.Text)
                Case 5
                    ValorNumero29.Text = Pusing("###,###.#####", ValorNumero29.Text)
                Case 6
                    ValorNumero29.Text = Pusing("###,###.######", ValorNumero29.Text)
                Case Else
                    ValorNumero29.Text = Pusing("###,###", ValorNumero29.Text)
            End Select
            
            Valor29.Text = ValorNumero29.Text + " " + ZUnidad(29)
            
            ValorNumero30.SetFocus
            
                Else
                
            If ValorNumero29.Text = "S" Or ValorNumero29.Text = "N" Then
                If ValorNumero29.Text = "S" Then
                    Valor29.Text = "Cumple"
                        Else
                    Valor29.Text = "No Cumple"
                End If
                ValorNumero30.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero29.Text = ""
    End If
    
    If Val(ZDesde(29)) <> 0 Or Val(ZHasta(29)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub



Private Sub ValorNumero30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        If Val(ZDesde(30)) <> 0 Or Val(ZHasta(30)) <> 0 Then
        
            ZNumeI = InStr(Trim(ZDesde(30)), ".")
            ZNumeII = Len(Trim(ZDesde(30)))
            If ZNumeI <> 0 Then
                ZDife = ZNumeII - ZNumeI
                    Else
                ZDife = 0
            End If
            Select Case ZDife
                Case 1
                    ValorNumero30.Text = Pusing("###,###.#", ValorNumero30.Text)
                Case 2
                    ValorNumero30.Text = Pusing("###,###.##", ValorNumero30.Text)
                Case 3
                    ValorNumero30.Text = Pusing("###,###.###", ValorNumero30.Text)
                Case 4
                    ValorNumero30.Text = Pusing("###,###.####", ValorNumero30.Text)
                Case 5
                    ValorNumero30.Text = Pusing("###,###.#####", ValorNumero30.Text)
                Case 6
                    ValorNumero30.Text = Pusing("###,###.######", ValorNumero30.Text)
                Case Else
                    ValorNumero30.Text = Pusing("###,###", ValorNumero30.Text)
            End Select
            
            Valor30.Text = ValorNumero30.Text + " " + ZUnidad(30)
            
            SSTab1.Tab = 0
            ValorNumero1.SetFocus
            
                Else
                
            If ValorNumero30.Text = "S" Or ValorNumero30.Text = "N" Then
                If ValorNumero30.Text = "S" Then
                    Valor30.Text = "Cumple"
                        Else
                    Valor30.Text = "No Cumple"
                End If
                SSTab1.Tab = 0
                ValorNumero1.SetFocus
            End If
            
        End If
    End If
    
    If KeyAscii = 27 Then
        ValorNumero30.Text = ""
    End If
    
    If Val(ZDesde(30)) <> 0 Or Val(ZHasta(30)) <> 0 Then
        Call NumbersOnly(Screen.ActiveControl, KeyAscii)
            Else
        ZControl = Chr$(KeyAscii)
        If ZControl <> "S" And ZControl <> "N" Then
            KeyAscii = 0
        End If
    End If
    
End Sub








































Private Sub Ensayo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Aspecto.SetFocus
    End If
End Sub

Private Sub Aspecto_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Observaciones.SetFocus
    End If
End Sub

Private Sub Observaciones_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Confecciono.SetFocus
    End If
End Sub

Private Sub Confecciono_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.SetFocus
    End If
End Sub

Private Sub Lote_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spLaudo = "ListaLaudo " + "'" + Lote.Text + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            m$ = "Numero de lote ya existente"
            A% = MsgBox(m$, 0, "Pruebas de Materias Primas")
            rstLaudo.Close
                Else
            Liberada.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Liberada_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Liberada.Text = Pusing("###,###.##", Liberada.Text)
        Devuelta.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Devuelta_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Devuelta.Text = Pusing("###,###.##", Devuelta.Text)
        If Val(Devuelta.Text) = 0 Then
            NroRechazo.Text = ""
            Nueva.Text = "N"
            PartidaProveedor.SetFocus
                Else
            NroRechazo.SetFocus
        End If
    End If
End Sub

Private Sub NroRechazo_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spLaudo = "ListaLaudo " + "'" + NroRechazo.Text + "'"
        Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        If rstLaudo.RecordCount > 0 Then
            m$ = "Numero de lote ya existente"
            A% = MsgBox(m$, 0, "Pruebas de Materias Primas")
            rstLaudo.Close
                Else
            Nueva.SetFocus
        End If
    End If
End Sub

Private Sub Nueva_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Verifica_datos(Nueva.Text, "SN", Auxi4)
        If Auxi4 = "S" Then
            PartidaProveedor.SetFocus
        End If
    End If
End Sub

Private Sub PartidaProveedor_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OrigenMercaderia.SetFocus
    End If
End Sub

Private Sub OrigenMercaderia_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Liberada.SetFocus
    End If
End Sub

Private Sub imprime_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select



    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Producto.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        Producto.Text = rstEspecificacionesUnifica!Producto
        
        Rem Ensayo1.Caption = rstEspecificacionesUnifica!Ensayo1
        Rem Ensayo2.Caption = rstEspecificacionesUnifica!Ensayo2
        Rem Ensayo3.Caption = rstEspecificacionesUnifica!Ensayo3
        Rem Ensayo4.Caption = rstEspecificacionesUnifica!Ensayo4
        Rem Ensayo5.Caption = rstEspecificacionesUnifica!Ensayo5
        Rem Ensayo6.Caption = rstEspecificacionesUnifica!Ensayo6
        Rem Ensayo7.Caption = rstEspecificacionesUnifica!Ensayo7
        Rem Ensayo8.Caption = rstEspecificacionesUnifica!Ensayo8
        Rem Ensayo9.Caption = rstEspecificacionesUnifica!Ensayo9
        Rem Ensayo10.Caption = rstEspecificacionesUnifica!Ensayo10
        Rem Ensayo11.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        Rem Ensayo12.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        Rem Ensayo13.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        Rem Ensayo14.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        Rem Ensayo15.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        Rem Ensayo16.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        Rem Ensayo17.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        Rem Ensayo18.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        Rem Ensayo19.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        Rem Ensayo20.Caption = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        ZEnsayo1 = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo2 = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo3 = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo4 = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo5 = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo6 = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo7 = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo8 = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo9 = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo10 = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        Std1.Caption = rstEspecificacionesUnifica!Valor1
        Std2.Caption = rstEspecificacionesUnifica!Valor2
        Std3.Caption = rstEspecificacionesUnifica!Valor3
        Std4.Caption = rstEspecificacionesUnifica!Valor4
        Std5.Caption = rstEspecificacionesUnifica!Valor5
        Std6.Caption = rstEspecificacionesUnifica!Valor6
        Std7.Caption = rstEspecificacionesUnifica!Valor7
        Std8.Caption = rstEspecificacionesUnifica!Valor8
        Std9.Caption = rstEspecificacionesUnifica!Valor9
        Std10.Caption = rstEspecificacionesUnifica!Valor10
        Std11.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
        Std12.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
        Std13.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
        Std14.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
        Std15.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
        Std16.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
        Std17.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
        Std18.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
        Std19.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
        Std20.Caption = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
        
        ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        ZDesde(1) = IIf(IsNull(rstEspecificacionesUnifica!Desde1), "", rstEspecificacionesUnifica!Desde1)
        ZDesde(2) = IIf(IsNull(rstEspecificacionesUnifica!Desde2), "", rstEspecificacionesUnifica!Desde2)
        ZDesde(3) = IIf(IsNull(rstEspecificacionesUnifica!Desde3), "", rstEspecificacionesUnifica!Desde3)
        ZDesde(4) = IIf(IsNull(rstEspecificacionesUnifica!Desde4), "", rstEspecificacionesUnifica!Desde4)
        ZDesde(5) = IIf(IsNull(rstEspecificacionesUnifica!Desde5), "", rstEspecificacionesUnifica!Desde5)
        ZDesde(6) = IIf(IsNull(rstEspecificacionesUnifica!Desde6), "", rstEspecificacionesUnifica!Desde6)
        ZDesde(7) = IIf(IsNull(rstEspecificacionesUnifica!Desde7), "", rstEspecificacionesUnifica!Desde7)
        ZDesde(8) = IIf(IsNull(rstEspecificacionesUnifica!Desde8), "", rstEspecificacionesUnifica!Desde8)
        ZDesde(9) = IIf(IsNull(rstEspecificacionesUnifica!Desde9), "", rstEspecificacionesUnifica!Desde9)
        ZDesde(10) = IIf(IsNull(rstEspecificacionesUnifica!Desde10), "", rstEspecificacionesUnifica!Desde10)
        ZDesde(11) = IIf(IsNull(rstEspecificacionesUnifica!Desde11), "", rstEspecificacionesUnifica!Desde11)
        ZDesde(12) = IIf(IsNull(rstEspecificacionesUnifica!Desde12), "", rstEspecificacionesUnifica!Desde12)
        ZDesde(13) = IIf(IsNull(rstEspecificacionesUnifica!Desde13), "", rstEspecificacionesUnifica!Desde13)
        ZDesde(14) = IIf(IsNull(rstEspecificacionesUnifica!Desde14), "", rstEspecificacionesUnifica!Desde14)
        ZDesde(15) = IIf(IsNull(rstEspecificacionesUnifica!Desde15), "", rstEspecificacionesUnifica!Desde15)
        ZDesde(16) = IIf(IsNull(rstEspecificacionesUnifica!Desde16), "", rstEspecificacionesUnifica!Desde16)
        ZDesde(17) = IIf(IsNull(rstEspecificacionesUnifica!Desde17), "", rstEspecificacionesUnifica!Desde17)
        ZDesde(18) = IIf(IsNull(rstEspecificacionesUnifica!Desde18), "", rstEspecificacionesUnifica!Desde18)
        ZDesde(19) = IIf(IsNull(rstEspecificacionesUnifica!Desde19), "", rstEspecificacionesUnifica!Desde19)
        ZDesde(20) = IIf(IsNull(rstEspecificacionesUnifica!Desde20), "", rstEspecificacionesUnifica!Desde20)
        
        ZHasta(1) = IIf(IsNull(rstEspecificacionesUnifica!Hasta1), "", rstEspecificacionesUnifica!Hasta1)
        ZHasta(2) = IIf(IsNull(rstEspecificacionesUnifica!Hasta2), "", rstEspecificacionesUnifica!Hasta2)
        ZHasta(3) = IIf(IsNull(rstEspecificacionesUnifica!Hasta3), "", rstEspecificacionesUnifica!Hasta3)
        ZHasta(4) = IIf(IsNull(rstEspecificacionesUnifica!Hasta4), "", rstEspecificacionesUnifica!Hasta4)
        ZHasta(5) = IIf(IsNull(rstEspecificacionesUnifica!Hasta5), "", rstEspecificacionesUnifica!Hasta5)
        ZHasta(6) = IIf(IsNull(rstEspecificacionesUnifica!Hasta6), "", rstEspecificacionesUnifica!Hasta6)
        ZHasta(7) = IIf(IsNull(rstEspecificacionesUnifica!Hasta7), "", rstEspecificacionesUnifica!Hasta7)
        ZHasta(8) = IIf(IsNull(rstEspecificacionesUnifica!Hasta8), "", rstEspecificacionesUnifica!Hasta8)
        ZHasta(9) = IIf(IsNull(rstEspecificacionesUnifica!Hasta9), "", rstEspecificacionesUnifica!Hasta9)
        ZHasta(10) = IIf(IsNull(rstEspecificacionesUnifica!Hasta10), "", rstEspecificacionesUnifica!Hasta10)
        ZHasta(11) = IIf(IsNull(rstEspecificacionesUnifica!Hasta11), "", rstEspecificacionesUnifica!Hasta11)
        ZHasta(12) = IIf(IsNull(rstEspecificacionesUnifica!Hasta12), "", rstEspecificacionesUnifica!Hasta12)
        ZHasta(13) = IIf(IsNull(rstEspecificacionesUnifica!Hasta13), "", rstEspecificacionesUnifica!Hasta13)
        ZHasta(14) = IIf(IsNull(rstEspecificacionesUnifica!Hasta14), "", rstEspecificacionesUnifica!Hasta14)
        ZHasta(15) = IIf(IsNull(rstEspecificacionesUnifica!Hasta15), "", rstEspecificacionesUnifica!Hasta15)
        ZHasta(16) = IIf(IsNull(rstEspecificacionesUnifica!Hasta16), "", rstEspecificacionesUnifica!Hasta16)
        ZHasta(17) = IIf(IsNull(rstEspecificacionesUnifica!Hasta17), "", rstEspecificacionesUnifica!Hasta17)
        ZHasta(18) = IIf(IsNull(rstEspecificacionesUnifica!Hasta18), "", rstEspecificacionesUnifica!Hasta18)
        ZHasta(19) = IIf(IsNull(rstEspecificacionesUnifica!Hasta19), "", rstEspecificacionesUnifica!Hasta19)
        ZHasta(20) = IIf(IsNull(rstEspecificacionesUnifica!Hasta20), "", rstEspecificacionesUnifica!Hasta20)
        
        ZDesde(1) = Trim(ZDesde(1))
        ZDesde(2) = Trim(ZDesde(2))
        ZDesde(3) = Trim(ZDesde(3))
        ZDesde(4) = Trim(ZDesde(4))
        ZDesde(5) = Trim(ZDesde(5))
        ZDesde(6) = Trim(ZDesde(6))
        ZDesde(7) = Trim(ZDesde(7))
        ZDesde(8) = Trim(ZDesde(8))
        ZDesde(9) = Trim(ZDesde(9))
        ZDesde(10) = Trim(ZDesde(10))
        ZDesde(11) = Trim(ZDesde(11))
        ZDesde(12) = Trim(ZDesde(12))
        ZDesde(13) = Trim(ZDesde(13))
        ZDesde(14) = Trim(ZDesde(14))
        ZDesde(15) = Trim(ZDesde(15))
        ZDesde(16) = Trim(ZDesde(16))
        ZDesde(17) = Trim(ZDesde(17))
        ZDesde(18) = Trim(ZDesde(18))
        ZDesde(19) = Trim(ZDesde(19))
        ZDesde(20) = Trim(ZDesde(20))
        
        ZHasta(1) = Trim(ZHasta(1))
        ZHasta(2) = Trim(ZHasta(2))
        ZHasta(3) = Trim(ZHasta(3))
        ZHasta(4) = Trim(ZHasta(4))
        ZHasta(5) = Trim(ZHasta(5))
        ZHasta(6) = Trim(ZHasta(6))
        ZHasta(7) = Trim(ZHasta(7))
        ZHasta(8) = Trim(ZHasta(8))
        ZHasta(9) = Trim(ZHasta(9))
        ZHasta(10) = Trim(ZHasta(10))
        ZHasta(11) = Trim(ZHasta(11))
        ZHasta(12) = Trim(ZHasta(12))
        ZHasta(13) = Trim(ZHasta(13))
        ZHasta(14) = Trim(ZHasta(14))
        ZHasta(15) = Trim(ZHasta(15))
        ZHasta(16) = Trim(ZHasta(16))
        ZHasta(17) = Trim(ZHasta(17))
        ZHasta(18) = Trim(ZHasta(18))
        ZHasta(19) = Trim(ZHasta(19))
        ZHasta(20) = Trim(ZHasta(20))
        
        rstEspecificacionesUnifica.Close
                        
    End If
    
    


    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Producto.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        ZEnsayo21 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo22 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo23 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo24 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo25 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo26 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo27 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo28 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo29 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo30 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        Std21.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
        Std22.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
        Std23.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
        Std24.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
        Std25.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
        Std26.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
        Std27.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
        Std28.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
        Std29.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
        Std30.Caption = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
        
        ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        ZDesde(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde21), "", rstEspecificacionesUnificaIII!Desde21)
        ZDesde(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde22), "", rstEspecificacionesUnificaIII!Desde22)
        ZDesde(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde23), "", rstEspecificacionesUnificaIII!Desde23)
        ZDesde(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde24), "", rstEspecificacionesUnificaIII!Desde24)
        ZDesde(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde25), "", rstEspecificacionesUnificaIII!Desde25)
        ZDesde(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde26), "", rstEspecificacionesUnificaIII!Desde26)
        ZDesde(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde27), "", rstEspecificacionesUnificaIII!Desde27)
        ZDesde(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde28), "", rstEspecificacionesUnificaIII!Desde28)
        ZDesde(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde29), "", rstEspecificacionesUnificaIII!Desde29)
        ZDesde(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Desde30), "", rstEspecificacionesUnificaIII!Desde30)
        
        ZHasta(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta21), "", rstEspecificacionesUnificaIII!Hasta21)
        ZHasta(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta22), "", rstEspecificacionesUnificaIII!Hasta22)
        ZHasta(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta23), "", rstEspecificacionesUnificaIII!Hasta23)
        ZHasta(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta24), "", rstEspecificacionesUnificaIII!Hasta24)
        ZHasta(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta25), "", rstEspecificacionesUnificaIII!Hasta25)
        ZHasta(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta26), "", rstEspecificacionesUnificaIII!Hasta26)
        ZHasta(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta27), "", rstEspecificacionesUnificaIII!Hasta27)
        ZHasta(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta28), "", rstEspecificacionesUnificaIII!Hasta28)
        ZHasta(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta29), "", rstEspecificacionesUnificaIII!Hasta29)
        ZHasta(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta30), "", rstEspecificacionesUnificaIII!Hasta30)
        
        ZDesde(21) = Trim(ZDesde(21))
        ZDesde(22) = Trim(ZDesde(22))
        ZDesde(23) = Trim(ZDesde(23))
        ZDesde(24) = Trim(ZDesde(24))
        ZDesde(25) = Trim(ZDesde(25))
        ZDesde(26) = Trim(ZDesde(26))
        ZDesde(27) = Trim(ZDesde(27))
        ZDesde(28) = Trim(ZDesde(28))
        ZDesde(29) = Trim(ZDesde(29))
        ZDesde(30) = Trim(ZDesde(30))
        
        ZHasta(21) = Trim(ZHasta(21))
        ZHasta(22) = Trim(ZHasta(22))
        ZHasta(23) = Trim(ZHasta(23))
        ZHasta(24) = Trim(ZHasta(24))
        ZHasta(25) = Trim(ZHasta(25))
        ZHasta(26) = Trim(ZHasta(26))
        ZHasta(27) = Trim(ZHasta(27))
        ZHasta(28) = Trim(ZHasta(28))
        ZHasta(29) = Trim(ZHasta(29))
        ZHasta(30) = Trim(ZHasta(30))
        
        rstEspecificacionesUnificaIII.Close
                        
    End If
    
    
    
    
    If Trim(Std1.Caption) = "" Then
        Rem Ensayo1.Caption = ""
        ZEnsayo1 = ""
        ZEnsayo(1) = 0
    End If
    
    If Trim(Std2.Caption) = "" Then
        Rem Ensayo2.Caption = ""
        ZEnsayo2 = ""
        ZEnsayo(2) = 0
    End If
    
    If Trim(Std3.Caption) = "" Then
        Rem Ensayo3.Caption = ""
        ZEnsayo3 = ""
        ZEnsayo(3) = 0
    End If
    
    If Trim(Std4.Caption) = "" Then
        Rem Ensayo4.Caption = ""
        ZEnsayo4 = ""
        ZEnsayo(4) = 0
    End If
    
    If Trim(Std5.Caption) = "" Then
        Rem Ensayo5.Caption = ""
        ZEnsayo5 = ""
        ZEnsayo(5) = 0
    End If
    
    If Trim(Std6.Caption) = "" Then
        Rem Ensayo6.Caption = ""
        ZEnsayo6 = ""
        ZEnsayo(6) = 0
    End If
    
    If Trim(Std7.Caption) = "" Then
        Rem Ensayo7.Caption = ""
        ZEnsayo7 = ""
        ZEnsayo(7) = 0
    End If
    
    If Trim(Std8.Caption) = "" Then
        Rem Ensayo8.Caption = ""
        ZEnsayo8 = ""
        ZEnsayo(8) = 0
    End If
    
    If Trim(Std9.Caption) = "" Then
        Rem Ensayo9.Caption = ""
        ZEnsayo9 = ""
        ZEnsayo(9) = 0
    End If
    
    If Trim(Std10.Caption) = "" Then
        Rem Ensayo10.Caption = ""
        ZEnsayo10 = ""
        ZEnsayo(10) = 0
    End If
    
    If Trim(Std11.Caption) = "" Then
        Rem Ensayo11.Caption = ""
        ZEnsayo11 = ""
        ZEnsayo(11) = 0
    End If
    
    If Trim(Std12.Caption) = "" Then
        Rem Ensayo12.Caption = ""
        ZEnsayo12 = ""
        ZEnsayo(12) = 0
    End If
    
    If Trim(Std13.Caption) = "" Then
        Rem Ensayo13.Caption = ""
        ZEnsayo13 = ""
        ZEnsayo(13) = 0
    End If
    
    If Trim(Std14.Caption) = "" Then
        Rem Ensayo14.Caption = ""
        ZEnsayo14 = ""
        ZEnsayo(14) = 0
    End If
    
    If Trim(Std15.Caption) = "" Then
        Rem Ensayo15.Caption = ""
        ZEnsayo15 = ""
        ZEnsayo(15) = 0
    End If
    
    If Trim(Std16.Caption) = "" Then
        Rem Ensayo16.Caption = ""
        ZEnsayo16 = ""
        ZEnsayo(16) = 0
    End If
    
    If Trim(Std17.Caption) = "" Then
        Rem Ensayo17.Caption = ""
        ZEnsayo17 = ""
        ZEnsayo(17) = 0
    End If
    
    If Trim(Std18.Caption) = "" Then
        Rem Ensayo18.Caption = ""
        ZEnsayo18 = ""
        ZEnsayo(18) = 0
    End If
    
    If Trim(Std19.Caption) = "" Then
        Rem Ensayo19.Caption = ""
        ZEnsayo19 = ""
        ZEnsayo(19) = 0
    End If
    
    If Trim(Std20.Caption) = "" Then
        Rem Ensayo20.Caption = ""
        ZEnsayo20 = ""
        ZEnsayo(20) = 0
    End If
    
    If Trim(Std21.Caption) = "" Then
        Rem Ensayo21.Caption = ""
        ZEnsayo21 = ""
        ZEnsayo(21) = 0
    End If
    
    If Trim(Std22.Caption) = "" Then
        Rem Ensayo22.Caption = ""
        ZEnsayo22 = ""
        ZEnsayo(22) = 0
    End If
    
    If Trim(Std23.Caption) = "" Then
        Rem Ensayo23.Caption = ""
        ZEnsayo23 = ""
        ZEnsayo(23) = 0
    End If
    
    If Trim(Std24.Caption) = "" Then
        Rem Ensayo24.Caption = ""
        ZEnsayo24 = ""
        ZEnsayo(24) = 0
    End If
    
    If Trim(Std25.Caption) = "" Then
        Rem Ensayo25.Caption = ""
        ZEnsayo25 = ""
        ZEnsayo(25) = 0
    End If
    
    If Trim(Std26.Caption) = "" Then
        Rem Ensayo26.Caption = ""
        ZEnsayo26 = ""
        ZEnsayo(26) = 0
    End If
    
    If Trim(Std27.Caption) = "" Then
        Rem Ensayo27.Caption = ""
        ZEnsayo27 = ""
        ZEnsayo(27) = 0
    End If
    
    If Trim(Std28.Caption) = "" Then
        Rem Ensayo28.Caption = ""
        ZEnsayo28 = ""
        ZEnsayo(28) = 0
    End If
    
    If Trim(Std29.Caption) = "" Then
        Rem Ensayo29.Caption = ""
        ZEnsayo29 = ""
        ZEnsayo(29) = 0
    End If
    
    If Trim(Std30.Caption) = "" Then
        Rem Ensayo30.Caption = ""
        ZEnsayo30 = ""
        ZEnsayo(30) = 0
    End If

    Call ImprimeII_Click
    
    Call Conecta_Empresa

End Sub

Private Sub ImprimeII_Click()

    Erase ZUnidad
    
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(1) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri1.Caption = rstEnsayo!Descripcion
        ZUnidad(1) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri1.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(2) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri2.Caption = rstEnsayo!Descripcion
        ZUnidad(2) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri2.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(3) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri3.Caption = rstEnsayo!Descripcion
        ZUnidad(3) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri3.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(4) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri4.Caption = rstEnsayo!Descripcion
        ZUnidad(4) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri4.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(5) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri5.Caption = rstEnsayo!Descripcion
        ZUnidad(5) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri5.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(6) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri6.Caption = rstEnsayo!Descripcion
        ZUnidad(6) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri6.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(7) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri7.Caption = rstEnsayo!Descripcion
        ZUnidad(7) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri7.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(8) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri8.Caption = rstEnsayo!Descripcion
        ZUnidad(8) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri8.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(9) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri9.Caption = rstEnsayo!Descripcion
        ZUnidad(9) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri9.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(10) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri10.Caption = rstEnsayo!Descripcion
        ZUnidad(10) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri10.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(11) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri11.Caption = rstEnsayo!Descripcion
        ZUnidad(11) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri11.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(12) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri12.Caption = rstEnsayo!Descripcion
        ZUnidad(12) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri12.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(13) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri13.Caption = rstEnsayo!Descripcion
        ZUnidad(13) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri13.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(14) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri14.Caption = rstEnsayo!Descripcion
        ZUnidad(14) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri14.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(15) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri15.Caption = rstEnsayo!Descripcion
        ZUnidad(15) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri15.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(16) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri16.Caption = rstEnsayo!Descripcion
        ZUnidad(16) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri16.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(17) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri17.Caption = rstEnsayo!Descripcion
        ZUnidad(17) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri17.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(18) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri18.Caption = rstEnsayo!Descripcion
        ZUnidad(18) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri18.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(19) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri19.Caption = rstEnsayo!Descripcion
        ZUnidad(19) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri19.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(20) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri20.Caption = rstEnsayo!Descripcion
        ZUnidad(20) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri20.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(21) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri21.Caption = rstEnsayo!Descripcion
        ZUnidad(21) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri21.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(22) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri22.Caption = rstEnsayo!Descripcion
        ZUnidad(22) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri22.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(23) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri23.Caption = rstEnsayo!Descripcion
        ZUnidad(23) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri23.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(24) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri24.Caption = rstEnsayo!Descripcion
        ZUnidad(24) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri24.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(25) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri25.Caption = rstEnsayo!Descripcion
        ZUnidad(25) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri25.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(26) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri26.Caption = rstEnsayo!Descripcion
        ZUnidad(26) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri26.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(27) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri27.Caption = rstEnsayo!Descripcion
        ZUnidad(27) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri27.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(28) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri28.Caption = rstEnsayo!Descripcion
        ZUnidad(28) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri28.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(29) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri29.Caption = rstEnsayo!Descripcion
        ZUnidad(29) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri29.Caption = ""
    End If
        
    spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(30) + "'"
    Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
    If rstEnsayo.RecordCount > 0 Then
        Descri30.Caption = rstEnsayo!Descripcion
        ZUnidad(30) = IIf(IsNull(rstEnsayo!Unidad), "", rstEnsayo!Unidad)
        rstEnsayo.Close
            Else
        Descri30.Caption = ""
    End If

End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "" Then
            Producto.Text = UCase(Producto.Text)
            
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
            
            Sql1 = "Select *"
            Sql2 = " FROM EspecificacionesUnifica"
            Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Producto.Text + "'"
            spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnifica.RecordCount > 0 Then
                rstEspecificacionesUnifica.Close
                Call Conecta_Empresa
                Call imprime_Click
                    Else
                Call Conecta_Empresa
                WProducto = Producto.Text
                CmdLimpiar_Click
                Producto.Text = WProducto
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                Producto.SetFocus
                Exit Sub
            End If
            
        End If
        fecha.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

    WPantalla.Visible = False
    Muestra.Visible = False
    pantalla.Visible = False
    Ayuda.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False

    Opcion.Clear
    
    Opcion.AddItem "Articulos"
    Opcion.AddItem "Pruebas"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    WPantalla.Visible = False
    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 1
            Call Limpia_Vector
            LugarVector = 0
            spPrueart = "ListaPruebaConsulta"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
            
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueart!Producto <> "" Then
                            If rstPrueart!Producto <> "  -     -   " Then
                                If rstPrueart!Producto <> Space$(10) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueart!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                            Else
                                        Muestra.TextMatrix(LugarVector, 1) = ""
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Mid$(rstPrueart!Prueba, 2, 6)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueart!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueart!fecha
                                    IngresaItem = rstPrueart!Prueba
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            
            End If
            
        Case Else
    End Select
            
    If XIndice = 0 Then
        pantalla.Visible = True
            Else
        Muestra.Visible = True
    End If
    
End Sub

Private Sub NumeroPrueba_Keypress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    
    WIndice.Clear
    PantaNumeroPrueba.Visible = False
    WPantalla.Visible = False
    Call Limpia_Vector
    LugarVector = 0
    
    Sql1 = "Select *"
    Sql2 = " FROM PrueArt"
    Sql3 = " WHERE Lote = " + "'" + NumeroPrueba.Text + "'"
    spPrueart = Sql1 + Sql2 + Sql3
    Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrueart.RecordCount > 0 Then
        If rstPrueart!Producto <> "" Then
            If rstPrueart!Producto <> "  -     -   " Then
                If rstPrueart!Producto <> Space$(10) Then
                    LugarVector = LugarVector + 1
                    If Left$(rstPrueart!Prueba, 1) = "1" Then
                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                    End If
                    Muestra.TextMatrix(LugarVector, 2) = Mid$(rstPrueart!Prueba, 2, 6)
                    Muestra.TextMatrix(LugarVector, 3) = rstPrueart!Producto
                    Muestra.TextMatrix(LugarVector, 4) = rstPrueart!fecha
                    IngresaItem = rstPrueart!Prueba
                    WIndice.AddItem IngresaItem
                End If
            End If
        End If
        rstPrueart.Close
    End If
    
    Muestra.TopRow = 1
    Muestra.Row = 1
    Muestra.Col = 1
    
    End If

End Sub



Private Sub WTitulo_dblClick(Index As Integer)

    If Index = 2 Then
        PantaNumeroPrueba.Height = 855
        PantaNumeroPrueba.Left = 3600
        PantaNumeroPrueba.Top = 6020
        PantaNumeroPrueba.Width = 4095
        PantaNumeroPrueba.Visible = True
        NumeroPrueba.Text = ""
        NumeroPrueba.SetFocus
    End If
    
    If Index = 3 Then
        ColumnaOpcion = 3
        Call Busqueda
    End If
    
    If Index = 4 Then
        ColumnaOpcion = 4
        Call Busqueda
    End If
    
End Sub

Private Sub Busqueda()

    WPantalla.Clear
    Select Case ColumnaOpcion
        Case 3
            WPantalla.AddItem ""
            Sql1 = "Select DISTINCT Producto"
            Sql2 = " FROM Prueart"
            Sql3 = " Order by Producto"
            spPrueart = Sql1 + Sql2 + Sql3
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueart!Producto <> "" Then
                            If rstPrueart!Producto <> "  -     -   " Then
                                If rstPrueart!Producto <> Space$(10) Then
                                    IngresaItem = rstPrueart!Producto
                                    WPantalla.AddItem IngresaItem
                                End If
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            
        Case 4
            WPantalla.AddItem ""
            Sql1 = "Select DISTINCT FechaOrd"
            Sql2 = " FROM Prueart"
            Sql3 = " Order by FechaOrd"
            spPrueart = Sql1 + Sql2 + Sql3
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Right$(rstPrueart!FechaOrd, 2) + "/" + Mid$(rstPrueart!FechaOrd, 5, 2) + "/" + Left$(rstPrueart!FechaOrd, 4)
                        WPantalla.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            
            
        Case Else
        
    End Select
    
    WPantalla.Height = 2205
    WPantalla.Left = 4800
    WPantalla.Top = 5520
    WPantalla.Width = 2755
    
    WPantalla.Visible = True
    
End Sub

Private Sub WPantalla_Click()

    If WPantalla.ListIndex <> 0 Then
        Seleccion = WPantalla.Text
            Else
        Seleccion = ""
        ColumnaOpcion = 0
    End If
    WPantalla.Visible = False
    WIndice.Clear
    
    Select Case ColumnaOpcion
        Case 0
            WPantalla.Visible = False
            Call Limpia_Vector
            LugarVector = 0
        
            spPrueart = "ListaPruebaConsulta"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
                With rstPrueart
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstPrueart!Producto <> "" Then
                                If rstPrueart!Producto <> "  -     -   " Then
                                    If rstPrueart!Producto <> Space$(10) Then
                                        LugarVector = LugarVector + 1
                                        If Left$(rstPrueart!Prueba, 1) = "1" Then
                                            Muestra.TextMatrix(LugarVector, 1) = "OK"
                                        End If
                                        Muestra.TextMatrix(LugarVector, 2) = Mid$(rstPrueart!Prueba, 2, 6)
                                        Muestra.TextMatrix(LugarVector, 3) = rstPrueart!Producto
                                        Muestra.TextMatrix(LugarVector, 4) = rstPrueart!fecha
                                        IngresaItem = rstPrueart!Prueba
                                        WIndice.AddItem IngresaItem
                                    End If
                                End If
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPrueart.Close
            End If
            
        Case 3
            Call Limpia_Vector
            LugarVector = 0
            
            Sql1 = "Select *"
            Sql2 = " FROM Prueart"
            Sql3 = " Where Producto = " + "'" + Seleccion + "'"
            Sql4 = " Order by Producto, Fechaord"
            spPrueart = Sql1 + Sql2 + Sql3 + Sql4
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueart!Producto <> "" Then
                            If rstPrueart!Producto <> "  -     -   " Then
                                If rstPrueart!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueart!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Mid$(rstPrueart!Prueba, 2, 6)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueart!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueart!fecha
                                    IngresaItem = rstPrueart!Prueba
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            Muestra.TopRow = 1
            Muestra.Row = 1
            Muestra.Col = 1
            
            End If
    
        Case 4
            Call Limpia_Vector
            LugarVector = 0
            
            Sql1 = "Select *"
            Sql2 = " FROM prueart"
            Sql3 = " Where Fecha = " + "'" + Seleccion + "'"
            Sql4 = " Order by Producto, Fechaord"
            spPrueart = Sql1 + Sql2 + Sql3 + Sql4
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
            With rstPrueart
                .MoveFirst
                Do
                    If .EOF = False Then
                        If rstPrueart!Producto <> "" Then
                            If rstPrueart!Producto <> "  -     -   " Then
                                If rstPrueart!Producto <> Space$(12) Then
                                    LugarVector = LugarVector + 1
                                    If Left$(rstPrueart!Prueba, 1) = "1" Then
                                        Muestra.TextMatrix(LugarVector, 1) = "OK"
                                    End If
                                    Muestra.TextMatrix(LugarVector, 2) = Mid$(rstPrueart!Prueba, 2, 6)
                                    Muestra.TextMatrix(LugarVector, 3) = rstPrueart!Producto
                                    Muestra.TextMatrix(LugarVector, 4) = rstPrueart!fecha
                                    IngresaItem = rstPrueart!Prueba
                                    WIndice.AddItem IngresaItem
                                End If
                            End If
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstPrueart.Close
            Muestra.TopRow = 1
            Muestra.Row = 1
            Muestra.Col = 1
            
            End If
            
        Case Else
        
    End Select
    
End Sub

Private Sub Muestra_Click()

    If Muestra.Row <> 0 Then

    Muestra.Visible = False
    WTitulo(1).Visible = False
    WTitulo(2).Visible = False
    WTitulo(3).Visible = False
    WTitulo(4).Visible = False
    Select Case XIndice
        Case 1
            Indice = Muestra.Row - 1
            ClavePrue$ = WIndice.List(Indice)
            spPrueart = "ConsultaPrueart" + "'" + ClavePrue$ + "'"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
                Partida.Text = Mid$(ClavePrue$, 2, 6)
                Producto.Text = rstPrueart!Producto
                fecha.Text = rstPrueart!fecha
                WFechaord = Right$(fecha.Text, 4) + Mid$(fecha.Text, 4, 2) + Left$(fecha.Text, 2)
                Orden.Text = rstPrueart!Orden
                
                Valor1.Text = rstPrueart!Valor1
                Valor2.Text = rstPrueart!Valor2
                Valor3.Text = rstPrueart!Valor3
                Valor4.Text = rstPrueart!Valor4
                Valor5.Text = rstPrueart!Valor5
                Valor6.Text = rstPrueart!Valor6
                Valor7.Text = rstPrueart!Valor7
                Valor8.Text = rstPrueart!Valor8
                Valor9.Text = rstPrueart!Valor9
                Valor10.Text = rstPrueart!Valor10
                Valor11.Text = IIf(IsNull(rstPrueart!Valor11), "", rstPrueart!Valor11)
                Valor12.Text = IIf(IsNull(rstPrueart!Valor12), "", rstPrueart!Valor12)
                Valor13.Text = IIf(IsNull(rstPrueart!Valor13), "", rstPrueart!Valor13)
                Valor14.Text = IIf(IsNull(rstPrueart!Valor14), "", rstPrueart!Valor14)
                Valor15.Text = IIf(IsNull(rstPrueart!Valor15), "", rstPrueart!Valor15)
                Valor16.Text = IIf(IsNull(rstPrueart!Valor16), "", rstPrueart!Valor16)
                Valor17.Text = IIf(IsNull(rstPrueart!Valor17), "", rstPrueart!Valor17)
                Valor18.Text = IIf(IsNull(rstPrueart!Valor18), "", rstPrueart!Valor18)
                Valor19.Text = IIf(IsNull(rstPrueart!Valor19), "", rstPrueart!Valor19)
                Valor20.Text = IIf(IsNull(rstPrueart!Valor20), "", rstPrueart!Valor20)
                Valor21.Text = IIf(IsNull(rstPrueart!Valor21), "", rstPrueart!Valor21)
                Valor22.Text = IIf(IsNull(rstPrueart!Valor22), "", rstPrueart!Valor22)
                Valor23.Text = IIf(IsNull(rstPrueart!Valor23), "", rstPrueart!Valor23)
                Valor24.Text = IIf(IsNull(rstPrueart!Valor24), "", rstPrueart!Valor24)
                Valor25.Text = IIf(IsNull(rstPrueart!Valor25), "", rstPrueart!Valor25)
                Valor26.Text = IIf(IsNull(rstPrueart!Valor26), "", rstPrueart!Valor26)
                Valor27.Text = IIf(IsNull(rstPrueart!Valor27), "", rstPrueart!Valor27)
                Valor28.Text = IIf(IsNull(rstPrueart!Valor28), "", rstPrueart!Valor28)
                Valor29.Text = IIf(IsNull(rstPrueart!Valor29), "", rstPrueart!Valor29)
                Valor30.Text = IIf(IsNull(rstPrueart!Valor30), "", rstPrueart!Valor30)
                
                ValorNumero1.Text = IIf(IsNull(rstPrueart!ValorNumero1), "", rstPrueart!ValorNumero1)
                ValorNumero2.Text = IIf(IsNull(rstPrueart!ValorNumero2), "", rstPrueart!ValorNumero2)
                ValorNumero3.Text = IIf(IsNull(rstPrueart!ValorNumero3), "", rstPrueart!ValorNumero3)
                ValorNumero4.Text = IIf(IsNull(rstPrueart!ValorNumero4), "", rstPrueart!ValorNumero4)
                ValorNumero5.Text = IIf(IsNull(rstPrueart!ValorNumero5), "", rstPrueart!ValorNumero5)
                ValorNumero6.Text = IIf(IsNull(rstPrueart!ValorNumero6), "", rstPrueart!ValorNumero6)
                ValorNumero7.Text = IIf(IsNull(rstPrueart!ValorNumero7), "", rstPrueart!ValorNumero7)
                ValorNumero8.Text = IIf(IsNull(rstPrueart!ValorNumero8), "", rstPrueart!ValorNumero8)
                ValorNumero9.Text = IIf(IsNull(rstPrueart!ValorNumero9), "", rstPrueart!ValorNumero9)
                ValorNumero10.Text = IIf(IsNull(rstPrueart!ValorNumero10), "", rstPrueart!ValorNumero10)
                ValorNumero11.Text = IIf(IsNull(rstPrueart!ValorNumero11), "", rstPrueart!ValorNumero11)
                ValorNumero12.Text = IIf(IsNull(rstPrueart!ValorNumero12), "", rstPrueart!ValorNumero12)
                ValorNumero13.Text = IIf(IsNull(rstPrueart!ValorNumero13), "", rstPrueart!ValorNumero13)
                ValorNumero14.Text = IIf(IsNull(rstPrueart!ValorNumero14), "", rstPrueart!ValorNumero14)
                ValorNumero15.Text = IIf(IsNull(rstPrueart!ValorNumero15), "", rstPrueart!ValorNumero15)
                ValorNumero16.Text = IIf(IsNull(rstPrueart!ValorNumero16), "", rstPrueart!ValorNumero16)
                ValorNumero17.Text = IIf(IsNull(rstPrueart!ValorNumero17), "", rstPrueart!ValorNumero17)
                ValorNumero18.Text = IIf(IsNull(rstPrueart!ValorNumero18), "", rstPrueart!ValorNumero18)
                ValorNumero19.Text = IIf(IsNull(rstPrueart!ValorNumero19), "", rstPrueart!ValorNumero19)
                ValorNumero20.Text = IIf(IsNull(rstPrueart!ValorNumero20), "", rstPrueart!ValorNumero20)
                ValorNumero21.Text = IIf(IsNull(rstPrueart!ValorNumero21), "", rstPrueart!ValorNumero21)
                ValorNumero22.Text = IIf(IsNull(rstPrueart!ValorNumero22), "", rstPrueart!ValorNumero22)
                ValorNumero23.Text = IIf(IsNull(rstPrueart!ValorNumero23), "", rstPrueart!ValorNumero23)
                ValorNumero24.Text = IIf(IsNull(rstPrueart!ValorNumero24), "", rstPrueart!ValorNumero24)
                ValorNumero25.Text = IIf(IsNull(rstPrueart!ValorNumero25), "", rstPrueart!ValorNumero25)
                ValorNumero26.Text = IIf(IsNull(rstPrueart!ValorNumero26), "", rstPrueart!ValorNumero26)
                ValorNumero27.Text = IIf(IsNull(rstPrueart!ValorNumero27), "", rstPrueart!ValorNumero27)
                ValorNumero28.Text = IIf(IsNull(rstPrueart!ValorNumero28), "", rstPrueart!ValorNumero28)
                ValorNumero29.Text = IIf(IsNull(rstPrueart!ValorNumero29), "", rstPrueart!ValorNumero29)
                ValorNumero30.Text = IIf(IsNull(rstPrueart!ValorNumero30), "", rstPrueart!ValorNumero30)
                
                ValorNumero1.Text = Trim(ValorNumero1.Text)
                ValorNumero2.Text = Trim(ValorNumero2.Text)
                ValorNumero3.Text = Trim(ValorNumero3.Text)
                ValorNumero4.Text = Trim(ValorNumero4.Text)
                ValorNumero5.Text = Trim(ValorNumero5.Text)
                ValorNumero6.Text = Trim(ValorNumero6.Text)
                ValorNumero7.Text = Trim(ValorNumero7.Text)
                ValorNumero8.Text = Trim(ValorNumero8.Text)
                ValorNumero9.Text = Trim(ValorNumero9.Text)
                ValorNumero10.Text = Trim(ValorNumero10.Text)
                ValorNumero11.Text = Trim(ValorNumero11.Text)
                ValorNumero12.Text = Trim(ValorNumero12.Text)
                ValorNumero13.Text = Trim(ValorNumero13.Text)
                ValorNumero14.Text = Trim(ValorNumero14.Text)
                ValorNumero15.Text = Trim(ValorNumero15.Text)
                ValorNumero16.Text = Trim(ValorNumero16.Text)
                ValorNumero17.Text = Trim(ValorNumero17.Text)
                ValorNumero18.Text = Trim(ValorNumero18.Text)
                ValorNumero19.Text = Trim(ValorNumero19.Text)
                ValorNumero20.Text = Trim(ValorNumero20.Text)
                ValorNumero21.Text = Trim(ValorNumero21.Text)
                ValorNumero22.Text = Trim(ValorNumero22.Text)
                ValorNumero23.Text = Trim(ValorNumero23.Text)
                ValorNumero24.Text = Trim(ValorNumero24.Text)
                ValorNumero25.Text = Trim(ValorNumero25.Text)
                ValorNumero26.Text = Trim(ValorNumero26.Text)
                ValorNumero27.Text = Trim(ValorNumero27.Text)
                ValorNumero28.Text = Trim(ValorNumero28.Text)
                ValorNumero29.Text = Trim(ValorNumero29.Text)
                ValorNumero30.Text = Trim(ValorNumero30.Text)
                
                Rem Descriprod.Caption = ""
                Rem Descri1.Caption = ""
                Ensayo.Text = rstPrueart!Ensayo
                Aspecto.Text = rstPrueart!Aspecto
                Observaciones.Text = rstPrueart!Observaciones
                Confecciono.Text = rstPrueart!Confecciono
                Rem Std1.Caption = ""
                Auxi = Left$(rstPrueart!Prueba, 1)
                Lote.Text = Right$(rstPrueart!Prueba, 6)
                Liberada.Text = rstPrueart!Liberada
                Devuelta.Text = rstPrueart!Devuelta
                NroRechazo.Text = rstPrueart!Rechazo
                Nueva.Text = rstPrueart!Nueva
                
                rstPrueart.Close
                
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
                
                LlamaImprime = "N"
                
                ZEnsayo1 = ""
                ZEnsayo2 = ""
                ZEnsayo3 = ""
                ZEnsayo4 = ""
                ZEnsayo5 = ""
                ZEnsayo6 = ""
                ZEnsayo7 = ""
                ZEnsayo8 = ""
                ZEnsayo9 = ""
                ZEnsayo10 = ""
                ZEnsayo11 = ""
                ZEnsayo12 = ""
                ZEnsayo13 = ""
                ZEnsayo14 = ""
                ZEnsayo15 = ""
                ZEnsayo16 = ""
                ZEnsayo17 = ""
                ZEnsayo18 = ""
                ZEnsayo19 = ""
                ZEnsayo20 = ""
                ZEnsayo21 = ""
                ZEnsayo22 = ""
                ZEnsayo23 = ""
                ZEnsayo24 = ""
                ZEnsayo25 = ""
                ZEnsayo26 = ""
                ZEnsayo27 = ""
                ZEnsayo28 = ""
                ZEnsayo29 = ""
                ZEnsayo30 = ""
                
                ZStd1 = ""
                ZStd2 = ""
                ZStd3 = ""
                ZStd4 = ""
                ZStd5 = ""
                ZStd6 = ""
                ZStd7 = ""
                ZStd8 = ""
                ZStd9 = ""
                ZStd10 = ""
                ZStd11 = ""
                ZStd12 = ""
                ZStd13 = ""
                ZStd14 = ""
                ZStd15 = ""
                ZStd16 = ""
                ZStd17 = ""
                ZStd18 = ""
                ZStd19 = ""
                ZStd20 = ""
                ZStd21 = ""
                ZStd22 = ""
                ZStd23 = ""
                ZStd24 = ""
                ZStd25 = ""
                ZStd26 = ""
                ZStd27 = ""
                ZStd28 = ""
                ZStd29 = ""
                ZStd30 = ""
                
                ZVersion = 0
                
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM EspecificacionesUnificaVersion"
                ZSql = ZSql + " Where EspecificacionesUnificaVersion.Producto = " + "'" + Producto.Text + "'"
                ZSql = ZSql + " Order by EspecificacionesUnificaVersion.Producto, EspecificacionesUnificaVersion.Version"
                
                spEspecificacionesUnificaVersion = ZSql
                Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnificaVersion.RecordCount > 0 Then
                    With rstEspecificacionesUnificaVersion
                        .MoveFirst
                        Do
                            If .EOF = False Then
                            
                                WDesde = Right$(rstEspecificacionesUnificaVersion!FechaInicio, 4) + Mid$(rstEspecificacionesUnificaVersion!FechaInicio, 4, 2) + Left$(rstEspecificacionesUnificaVersion!FechaInicio, 2)
                                WHasta = Right$(rstEspecificacionesUnificaVersion!FechaFinal, 4) + Mid$(rstEspecificacionesUnificaVersion!FechaFinal, 4, 2) + Left$(rstEspecificacionesUnificaVersion!FechaFinal, 2)
                                
                                If WDesde <= WFechaord And WHasta >= WFechaord Then
                                
                                    ZEnsayo1 = rstEspecificacionesUnificaVersion!Ensayo1
                                    ZEnsayo2 = rstEspecificacionesUnificaVersion!Ensayo2
                                    ZEnsayo3 = rstEspecificacionesUnificaVersion!Ensayo3
                                    ZEnsayo4 = rstEspecificacionesUnificaVersion!Ensayo4
                                    ZEnsayo5 = rstEspecificacionesUnificaVersion!Ensayo5
                                    ZEnsayo6 = rstEspecificacionesUnificaVersion!Ensayo6
                                    ZEnsayo7 = rstEspecificacionesUnificaVersion!Ensayo7
                                    ZEnsayo8 = rstEspecificacionesUnificaVersion!Ensayo8
                                    ZEnsayo9 = rstEspecificacionesUnificaVersion!Ensayo9
                                    ZEnsayo10 = rstEspecificacionesUnificaVersion!Ensayo10
                                    ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
                                    ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
                                    ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
                                    ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
                                    ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
                                    ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
                                    ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
                                    ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
                                    ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
                                    ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
                                    
                                    ZStd1 = rstEspecificacionesUnificaVersion!Valor1
                                    ZStd2 = rstEspecificacionesUnificaVersion!Valor2
                                    ZStd3 = rstEspecificacionesUnificaVersion!Valor3
                                    ZStd4 = rstEspecificacionesUnificaVersion!Valor4
                                    ZStd5 = rstEspecificacionesUnificaVersion!Valor5
                                    ZStd6 = rstEspecificacionesUnificaVersion!Valor6
                                    ZStd7 = rstEspecificacionesUnificaVersion!Valor7
                                    ZStd8 = rstEspecificacionesUnificaVersion!Valor8
                                    ZStd9 = rstEspecificacionesUnificaVersion!Valor9
                                    ZStd10 = rstEspecificacionesUnificaVersion!Valor10
                                    ZStd11 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor1), "", rstEspecificacionesUnificaVersion!ZValor1)
                                    ZStd12 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor2), "", rstEspecificacionesUnificaVersion!ZValor2)
                                    ZStd13 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor3), "", rstEspecificacionesUnificaVersion!ZValor3)
                                    ZStd14 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor4), "", rstEspecificacionesUnificaVersion!ZValor4)
                                    ZStd15 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor5), "", rstEspecificacionesUnificaVersion!ZValor5)
                                    ZStd16 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor6), "", rstEspecificacionesUnificaVersion!ZValor6)
                                    ZStd17 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor7), "", rstEspecificacionesUnificaVersion!ZValor7)
                                    ZStd18 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor8), "", rstEspecificacionesUnificaVersion!ZValor8)
                                    ZStd19 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor9), "", rstEspecificacionesUnificaVersion!ZValor9)
                                    ZStd20 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor10), "", rstEspecificacionesUnificaVersion!ZValor10)
                                    
                                    ZVersion = rstEspecificacionesUnificaVersion!Version
                                    ZZClave = rstEspecificacionesUnificaVersion!Clave
                                    LlamaImprime = "S"
                                End If
                                
                                If WDesde > WFechaord And LlamaImprime = "N" Then
                                
                                    ZEnsayo1 = rstEspecificacionesUnificaVersion!Ensayo1
                                    ZEnsayo2 = rstEspecificacionesUnificaVersion!Ensayo2
                                    ZEnsayo3 = rstEspecificacionesUnificaVersion!Ensayo3
                                    ZEnsayo4 = rstEspecificacionesUnificaVersion!Ensayo4
                                    ZEnsayo5 = rstEspecificacionesUnificaVersion!Ensayo5
                                    ZEnsayo6 = rstEspecificacionesUnificaVersion!Ensayo6
                                    ZEnsayo7 = rstEspecificacionesUnificaVersion!Ensayo7
                                    ZEnsayo8 = rstEspecificacionesUnificaVersion!Ensayo8
                                    ZEnsayo9 = rstEspecificacionesUnificaVersion!Ensayo9
                                    ZEnsayo10 = rstEspecificacionesUnificaVersion!Ensayo10
                                    ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo11), "", rstEspecificacionesUnificaVersion!Ensayo11)
                                    ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo12), "", rstEspecificacionesUnificaVersion!Ensayo12)
                                    ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo13), "", rstEspecificacionesUnificaVersion!Ensayo13)
                                    ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo14), "", rstEspecificacionesUnificaVersion!Ensayo14)
                                    ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo15), "", rstEspecificacionesUnificaVersion!Ensayo15)
                                    ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo16), "", rstEspecificacionesUnificaVersion!Ensayo16)
                                    ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo17), "", rstEspecificacionesUnificaVersion!Ensayo17)
                                    ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo18), "", rstEspecificacionesUnificaVersion!Ensayo18)
                                    ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo19), "", rstEspecificacionesUnificaVersion!Ensayo19)
                                    ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnificaVersion!Ensayo20), "", rstEspecificacionesUnificaVersion!Ensayo20)
                                    
                                    ZStd1 = rstEspecificacionesUnificaVersion!Valor1
                                    ZStd2 = rstEspecificacionesUnificaVersion!Valor2
                                    ZStd3 = rstEspecificacionesUnificaVersion!Valor3
                                    ZStd4 = rstEspecificacionesUnificaVersion!Valor4
                                    ZStd5 = rstEspecificacionesUnificaVersion!Valor5
                                    ZStd6 = rstEspecificacionesUnificaVersion!Valor6
                                    ZStd7 = rstEspecificacionesUnificaVersion!Valor7
                                    ZStd8 = rstEspecificacionesUnificaVersion!Valor8
                                    ZStd9 = rstEspecificacionesUnificaVersion!Valor9
                                    ZStd10 = rstEspecificacionesUnificaVersion!Valor10
                                    ZStd11 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor1), "", rstEspecificacionesUnificaVersion!ZValor1)
                                    ZStd12 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor2), "", rstEspecificacionesUnificaVersion!ZValor2)
                                    ZStd13 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor3), "", rstEspecificacionesUnificaVersion!ZValor3)
                                    ZStd14 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor4), "", rstEspecificacionesUnificaVersion!ZValor4)
                                    ZStd15 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor5), "", rstEspecificacionesUnificaVersion!ZValor5)
                                    ZStd16 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor6), "", rstEspecificacionesUnificaVersion!ZValor6)
                                    ZStd17 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor7), "", rstEspecificacionesUnificaVersion!ZValor7)
                                    ZStd18 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor8), "", rstEspecificacionesUnificaVersion!ZValor8)
                                    ZStd19 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor9), "", rstEspecificacionesUnificaVersion!ZValor9)
                                    ZStd20 = IIf(IsNull(rstEspecificacionesUnificaVersion!ZValor10), "", rstEspecificacionesUnificaVersion!ZValor10)
                                    
                                    ZVersion = rstEspecificacionesUnificaVersion!Version
                                    ZZClave = rstEspecificacionesUnificaVersion!Clave
                                    LlamaImprime = "S"
                                End If
                                
                                .MoveNext
                                    Else
                                Exit Do
                            End If
                        Loop
                    End With
                    rstEspecificacionesUnificaVersion.Close
                End If
                
                If LlamaImprime = "N" Then
                
                    Sql1 = "Select *"
                    Sql2 = " FROM EspecificacionesUnifica"
                    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Producto.Text + "'"
                    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
                    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspecificacionesUnifica.RecordCount > 0 Then
                    
                        ZEnsayo1 = rstEspecificacionesUnifica!Ensayo1
                        ZEnsayo2 = rstEspecificacionesUnifica!Ensayo2
                        ZEnsayo3 = rstEspecificacionesUnifica!Ensayo3
                        ZEnsayo4 = rstEspecificacionesUnifica!Ensayo4
                        ZEnsayo5 = rstEspecificacionesUnifica!Ensayo5
                        ZEnsayo6 = rstEspecificacionesUnifica!Ensayo6
                        ZEnsayo7 = rstEspecificacionesUnifica!Ensayo7
                        ZEnsayo8 = rstEspecificacionesUnifica!Ensayo8
                        ZEnsayo9 = rstEspecificacionesUnifica!Ensayo9
                        ZEnsayo10 = rstEspecificacionesUnifica!Ensayo10
                        ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
                        ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
                        ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
                        ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
                        ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
                        ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
                        ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
                        ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
                        ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
                        ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
                        
                        ZStd1 = rstEspecificacionesUnifica!Valor1
                        ZStd2 = rstEspecificacionesUnifica!Valor2
                        ZStd3 = rstEspecificacionesUnifica!Valor3
                        ZStd4 = rstEspecificacionesUnifica!Valor4
                        ZStd5 = rstEspecificacionesUnifica!Valor5
                        ZStd6 = rstEspecificacionesUnifica!Valor6
                        ZStd7 = rstEspecificacionesUnifica!Valor7
                        ZStd8 = rstEspecificacionesUnifica!Valor8
                        ZStd9 = rstEspecificacionesUnifica!Valor9
                        ZStd10 = rstEspecificacionesUnifica!Valor10
                        ZStd11 = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
                        ZStd12 = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
                        ZStd13 = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
                        ZStd14 = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
                        ZStd15 = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
                        ZStd16 = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
                        ZStd17 = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
                        ZStd18 = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
                        ZStd19 = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
                        ZStd20 = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
                        
                        ZVersion = rstEspecificacionesUnifica!Version
                        rstEspecificacionesUnifica.Close
                        LlamaImprime = "S"
                    End If
                
                
                    Sql1 = "Select *"
                    Sql2 = " FROM EspecificacionesUnificaIII"
                    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Producto.Text + "'"
                    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
                    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
                    
                        ZEnsayo21 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo21), "", rstEspecificacionesUnifica!Ensayo21)
                        ZEnsayo22 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo22), "", rstEspecificacionesUnifica!Ensayo22)
                        ZEnsayo23 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo23), "", rstEspecificacionesUnifica!Ensayo23)
                        ZEnsayo24 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo24), "", rstEspecificacionesUnifica!Ensayo24)
                        ZEnsayo25 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo25), "", rstEspecificacionesUnifica!Ensayo25)
                        ZEnsayo26 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo26), "", rstEspecificacionesUnifica!Ensayo26)
                        ZEnsayo27 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo27), "", rstEspecificacionesUnifica!Ensayo27)
                        ZEnsayo28 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo28), "", rstEspecificacionesUnifica!Ensayo28)
                        ZEnsayo29 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo29), "", rstEspecificacionesUnifica!Ensayo29)
                        ZEnsayo30 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo30), "", rstEspecificacionesUnifica!Ensayo30)
                        
                        ZStd21 = IIf(IsNull(rstEspecificacionesUnifica!Valor21), "", rstEspecificacionesUnifica!Valor21)
                        ZStd22 = IIf(IsNull(rstEspecificacionesUnifica!Valor22), "", rstEspecificacionesUnifica!Valor22)
                        ZStd23 = IIf(IsNull(rstEspecificacionesUnifica!Valor23), "", rstEspecificacionesUnifica!Valor23)
                        ZStd24 = IIf(IsNull(rstEspecificacionesUnifica!Valor24), "", rstEspecificacionesUnifica!Valor24)
                        ZStd25 = IIf(IsNull(rstEspecificacionesUnifica!Valor25), "", rstEspecificacionesUnifica!Valor25)
                        ZStd26 = IIf(IsNull(rstEspecificacionesUnifica!Valor26), "", rstEspecificacionesUnifica!Valor26)
                        ZStd27 = IIf(IsNull(rstEspecificacionesUnifica!Valor27), "", rstEspecificacionesUnifica!Valor27)
                        ZStd28 = IIf(IsNull(rstEspecificacionesUnifica!Valor28), "", rstEspecificacionesUnifica!Valor28)
                        ZStd29 = IIf(IsNull(rstEspecificacionesUnifica!Valor29), "", rstEspecificacionesUnifica!Valor29)
                        ZStd30 = IIf(IsNull(rstEspecificacionesUnifica!Valor30), "", rstEspecificacionesUnifica!Valor30)
                        
                        rstEspecificacionesUnificaIII.Close
                    End If
                
                
                
                End If
                
                If LlamaImprime = "S" Then
                
                
                    ZSql = ""
                    ZSql = ZSql + "Select *"
                    ZSql = ZSql + " FROM EspecificacionesUnificaVersionII"
                    ZSql = ZSql + " Where EspecificacionesUnificaVersionII.Clave = " + "'" + ZZClave + "'"
                    spEspecificacionesUnificaVersionII = ZSql
                    Set rstEspecificacionesUnificaVersionII = db.OpenRecordset(spEspecificacionesUnificaVersionII, dbOpenSnapshot, dbSQLPassThrough)
                    If rstEspecificacionesUnificaVersionII.RecordCount > 0 Then
                    
                        ZEnsayo21 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo21), "", rstEspecificacionesUnificaVersionII!Ensayo21)
                        ZEnsayo22 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo22), "", rstEspecificacionesUnificaVersionII!Ensayo22)
                        ZEnsayo23 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo23), "", rstEspecificacionesUnificaVersionII!Ensayo23)
                        ZEnsayo24 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo24), "", rstEspecificacionesUnificaVersionII!Ensayo24)
                        ZEnsayo25 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo25), "", rstEspecificacionesUnificaVersionII!Ensayo25)
                        ZEnsayo26 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo26), "", rstEspecificacionesUnificaVersionII!Ensayo26)
                        ZEnsayo27 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo27), "", rstEspecificacionesUnificaVersionII!Ensayo27)
                        ZEnsayo28 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo28), "", rstEspecificacionesUnificaVersionII!Ensayo28)
                        ZEnsayo29 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo29), "", rstEspecificacionesUnificaVersionII!Ensayo29)
                        ZEnsayo30 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Ensayo30), "", rstEspecificacionesUnificaVersionII!Ensayo30)
                        
                        ZStd21 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor21), "", rstEspecificacionesUnificaVersion!Valor21)
                        ZStd22 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor22), "", rstEspecificacionesUnificaVersion!Valor22)
                        ZStd23 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor23), "", rstEspecificacionesUnificaVersion!Valor23)
                        ZStd24 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor24), "", rstEspecificacionesUnificaVersion!Valor24)
                        ZStd25 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor25), "", rstEspecificacionesUnificaVersion!Valor25)
                        ZStd26 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor26), "", rstEspecificacionesUnificaVersion!Valor26)
                        ZStd27 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor27), "", rstEspecificacionesUnificaVersion!Valor27)
                        ZStd28 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor28), "", rstEspecificacionesUnificaVersion!Valor28)
                        ZStd29 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor29), "", rstEspecificacionesUnificaVersion!Valor29)
                        ZStd30 = IIf(IsNull(rstEspecificacionesUnificaVersionII!Valor30), "", rstEspecificacionesUnificaVersion!Valor30)
                                        
                        rstEspecificacionesUnificaVersionII.Close
                    End If
                
                    Rem Ensayo1.Caption = ZEnsayo1
                    Rem Ensayo2.Caption = ZEnsayo2
                    Rem Ensayo3.Caption = ZEnsayo3
                    Rem Ensayo4.Caption = ZEnsayo4
                    Rem Ensayo5.Caption = ZEnsayo5
                    Rem Ensayo6.Caption = ZEnsayo6
                    Rem Ensayo7.Caption = ZEnsayo7
                    Rem Ensayo8.Caption = ZEnsayo8
                    Rem Ensayo9.Caption = ZEnsayo9
                    Rem Ensayo10.Caption = ZEnsayo10
                    Rem Ensayo11.Caption = ZEnsayo11
                    Rem Ensayo12.Caption = ZEnsayo12
                    Rem Ensayo13.Caption = ZEnsayo13
                    Rem Ensayo14.Caption = ZEnsayo14
                    Rem Ensayo15.Caption = ZEnsayo15
                    Rem Ensayo16.Caption = ZEnsayo16
                    Rem Ensayo17.Caption = ZEnsayo17
                    Rem Ensayo18.Caption = ZEnsayo18
                    Rem Ensayo19.Caption = ZEnsayo19
                    Rem Ensayo20.Caption = ZEnsayo20
                    
                    ZEnsayo(1) = ZEnsayo1
                    ZEnsayo(2) = ZEnsayo2
                    ZEnsayo(3) = ZEnsayo3
                    ZEnsayo(4) = ZEnsayo4
                    ZEnsayo(5) = ZEnsayo5
                    ZEnsayo(6) = ZEnsayo6
                    ZEnsayo(7) = ZEnsayo7
                    ZEnsayo(8) = ZEnsayo8
                    ZEnsayo(9) = ZEnsayo9
                    ZEnsayo(10) = ZEnsayo10
                    ZEnsayo(11) = ZEnsayo11
                    ZEnsayo(12) = ZEnsayo12
                    ZEnsayo(13) = ZEnsayo13
                    ZEnsayo(14) = ZEnsayo14
                    ZEnsayo(15) = ZEnsayo15
                    ZEnsayo(16) = ZEnsayo16
                    ZEnsayo(17) = ZEnsayo17
                    ZEnsayo(18) = ZEnsayo18
                    ZEnsayo(19) = ZEnsayo19
                    ZEnsayo(20) = ZEnsayo20
                    ZEnsayo(21) = ZEnsayo21
                    ZEnsayo(22) = ZEnsayo22
                    ZEnsayo(23) = ZEnsayo23
                    ZEnsayo(24) = ZEnsayo24
                    ZEnsayo(25) = ZEnsayo25
                    ZEnsayo(26) = ZEnsayo26
                    ZEnsayo(27) = ZEnsayo27
                    ZEnsayo(28) = ZEnsayo28
                    ZEnsayo(29) = ZEnsayo29
                    ZEnsayo(30) = ZEnsayo30
                    
                    Std1.Caption = ZStd1
                    Std2.Caption = ZStd2
                    Std3.Caption = ZStd3
                    Std4.Caption = ZStd4
                    Std5.Caption = ZStd5
                    Std6.Caption = ZStd6
                    Std7.Caption = ZStd7
                    Std8.Caption = ZStd8
                    Std9.Caption = ZStd9
                    Std10.Caption = ZStd10
                    Std11.Caption = ZStd11
                    Std12.Caption = ZStd12
                    Std13.Caption = ZStd13
                    Std14.Caption = ZStd14
                    Std15.Caption = ZStd15
                    Std16.Caption = ZStd16
                    Std17.Caption = ZStd17
                    Std18.Caption = ZStd18
                    Std19.Caption = ZStd19
                    Std20.Caption = ZStd20
                    Std21.Caption = ZStd21
                    Std22.Caption = ZStd22
                    Std23.Caption = ZStd23
                    Std24.Caption = ZStd24
                    Std25.Caption = ZStd25
                    Std26.Caption = ZStd26
                    Std27.Caption = ZStd27
                    Std28.Caption = ZStd28
                    Std29.Caption = ZStd29
                    Std30.Caption = ZStd30
                    
                    If Trim(Std1.Caption) = "" Then
                        Rem Ensayo1.Caption = ""
                        ZEnsayo(1) = ""
                    End If
                    If Trim(Std2.Caption) = "" Then
                        Rem Ensayo2.Caption = ""
                        ZEnsayo(2) = ""
                    End If
                    If Trim(Std3.Caption) = "" Then
                        Rem Ensayo3.Caption = ""
                        ZEnsayo(3) = ""
                    End If
                    If Trim(Std4.Caption) = "" Then
                        Rem Ensayo4.Caption = ""
                        ZEnsayo(4) = ""
                    End If
                    If Trim(Std5.Caption) = "" Then
                        Rem Ensayo5.Caption = ""
                        ZEnsayo(5) = ""
                    End If
                    If Trim(Std6.Caption) = "" Then
                        Rem Ensayo6.Caption = ""
                        ZEnsayo(6) = ""
                    End If
                    If Trim(Std7.Caption) = "" Then
                        Rem Ensayo7.Caption = ""
                        ZEnsayo(7) = ""
                    End If
                    If Trim(Std8.Caption) = "" Then
                        Rem Ensayo8.Caption = ""
                        ZEnsayo(8) = ""
                    End If
                    If Trim(Std9.Caption) = "" Then
                        Rem Ensayo9.Caption = ""
                        ZEnsayo(9) = ""
                    End If
                    If Trim(Std10.Caption) = "" Then
                        Rem Ensayo10.Caption = ""
                        ZEnsayo(10) = ""
                    End If
                    If Trim(Std11.Caption) = "" Then
                        Rem Ensayo11.Caption = ""
                        ZEnsayo(11) = ""
                    End If
                    If Trim(Std12.Caption) = "" Then
                        Rem Ensayo12.Caption = ""
                        ZEnsayo(12) = ""
                    End If
                    If Trim(Std13.Caption) = "" Then
                        Rem Ensayo13.Caption = ""
                        ZEnsayo(13) = ""
                    End If
                    If Trim(Std14.Caption) = "" Then
                        Rem Ensayo14.Caption = ""
                        ZEnsayo(14) = ""
                    End If
                    If Trim(Std15.Caption) = "" Then
                        Rem Ensayo15.Caption = ""
                        ZEnsayo(15) = ""
                    End If
                    If Trim(Std16.Caption) = "" Then
                        Rem Ensayo16.Caption = ""
                        ZEnsayo(16) = ""
                    End If
                    If Trim(Std17.Caption) = "" Then
                        Rem Ensayo17.Caption = ""
                        ZEnsayo(17) = ""
                    End If
                    If Trim(Std18.Caption) = "" Then
                        Rem Ensayo18.Caption = ""
                        ZEnsayo(18) = ""
                    End If
                    If Trim(Std19.Caption) = "" Then
                        Rem Ensayo19.Caption = ""
                        ZEnsayo(19) = ""
                    End If
                    If Trim(Std20.Caption) = "" Then
                        Rem Ensayo20.Caption = ""
                        ZEnsayo(20) = ""
                    End If
                    If Trim(Std21.Caption) = "" Then
                        Rem Ensayo21.Caption = ""
                        ZEnsayo(21) = ""
                    End If
                    If Trim(Std22.Caption) = "" Then
                        Rem Ensayo22.Caption = ""
                        ZEnsayo(22) = ""
                    End If
                    If Trim(Std23.Caption) = "" Then
                        Rem Ensayo23.Caption = ""
                        ZEnsayo(23) = ""
                    End If
                    If Trim(Std24.Caption) = "" Then
                        Rem Ensayo24.Caption = ""
                        ZEnsayo(24) = ""
                    End If
                    If Trim(Std25.Caption) = "" Then
                        Rem Ensayo25.Caption = ""
                        ZEnsayo(25) = ""
                    End If
                    If Trim(Std26.Caption) = "" Then
                        Rem Ensayo26.Caption = ""
                        ZEnsayo(26) = ""
                    End If
                    If Trim(Std27.Caption) = "" Then
                        Rem Ensayo27.Caption = ""
                        ZEnsayo(27) = ""
                    End If
                    If Trim(Std28.Caption) = "" Then
                        Rem Ensayo28.Caption = ""
                        ZEnsayo(28) = ""
                    End If
                    If Trim(Std29.Caption) = "" Then
                        Rem Ensayo29.Caption = ""
                        ZEnsayo(29) = ""
                    End If
                    If Trim(Std30.Caption) = "" Then
                        Rem Ensayo30.Caption = ""
                        ZEnsayo(30) = ""
                    End If
                    
                    Call ImprimeII_Click
                End If
                        
                Call Conecta_Empresa
                
                spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Descriprod.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                spLaudo = "ListaLaudo " + "'" + Lote.Text + "'"
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    ZZFecha = rstLaudo!fecha
                    OrigenMercaderiaII.Text = rstLaudo!Origen
                    PartidaProveedorII.Text = rstLaudo!partiori
                    WVto = IIf(IsNull(rstLaudo!fechavencimiento), "  /  /    ", rstLaudo!fechavencimiento)
                    Vto.Text = WVto
                    WRevalida = IIf(IsNull(rstLaudo!Revalida), "0", rstLaudo!Revalida)
                    NroRevalida.Text = Str$(WRevalida)
                    Informe.Text = rstLaudo!Informe
                    rstLaudo.Close
                End If
                
                If Vto.Text = "  /  /    " Then
                
                    WVida = 0
                
                    spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WVida = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
                        rstArticulo.Close
                    End If
                        
                    WMes = Val(Mid$(ZZFecha, 4, 2))
                    WAno = Val(Right$(ZZFecha, 4))
                
                    For Ciclo = 1 To WVida
                        WMes = WMes + 1
                        If WMes > 12 Then
                            WAno = WAno + 1
                            WMes = 1
                        End If
                    Next Ciclo
                    If WVida <> 0 Then
                        XMes = Str$(WMes)
                        XAno = Str$(WAno)
                        Call Ceros(XMes, 2)
                        Call Ceros(XAno, 4)
                        Vto.Text = "01/" + XMes + "/" + XAno
                    End If
                
                End If
                
                lblresultado.Caption = "Valor Standard (Version : " + ZVersion + ")"
                lblresultadoII.Caption = "Valor Standard (Version : " + ZVersion + ")"
                    
                    Else
                    
                Call CmdLimpiar_Click
                
            End If
            Producto.SetFocus
        
        Case Else
    End Select
    
    End If

End Sub

Private Sub Limpia_Vector()

    Muestra.Clear
    
    Muestra.Height = 2775
    Muestra.Left = 2160
    Muestra.Top = 5350
    Muestra.Width = 7095

    Rem ponga la muestra en negritas
    Rem Muestra.Font.Bold = True

    ' Establesco loa Valores de la muestra
    
    Muestra.FixedCols = 1
    Muestra.Cols = 5
    Muestra.FixedRows = 1
    Muestra.Rows = 50000
    
    Muestra.ColWidth(0) = 200
    Muestra.Row = 0
    
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        Select Case Ciclo
            Case 1
                Muestra.Text = "Tipo"
                Muestra.ColWidth(Ciclo) = 1300
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 2
                Muestra.Text = "Nro.Prueba"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 3
                Muestra.Text = "Producto"
                Muestra.ColWidth(Ciclo) = 2000
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
            Case 4
                Muestra.Text = "Fecha"
                Muestra.ColWidth(Ciclo) = 1600
                Muestra.ColAlignment(Ciclo) = flexAlignLeftCenter
        End Select
    Next Ciclo
    
    
    Muestra.Row = 0
    For Ciclo = 1 To Muestra.Cols - 1
        Muestra.Col = Ciclo
        WTitulo(Ciclo).Text = Muestra.Text
        WTitulo(Ciclo).Left = Muestra.CellLeft + Muestra.Left
        WTitulo(Ciclo).Top = Muestra.CellTop + Muestra.Top
        WTitulo(Ciclo).Width = Muestra.CellWidth
        WTitulo(Ciclo).Height = Muestra.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA muestra
    
    WAncho = 340
    For Ciclo = 0 To Muestra.Cols - 1
        WAncho = WAncho + Muestra.ColWidth(Ciclo)
    Next Ciclo
    Muestra.Width = WAncho

    ' Size the columns.
    Font.Name = Muestra.Font.Name
    Font.Size = Muestra.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    Muestra.AllowUserResizing = flexResizeBoth
    
    Muestra.Col = 1
    Muestra.Row = 1
    
    
End Sub


Private Sub pantalla_Click()
    pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Clavepro$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Clavepro$ + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Producto.Text = rstArticulo!Codigo
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                Call imprime_Click
                    Else
                CmdLimpiar_Click
                Producto.Text = ""
                Descriprod.Caption = ""
            End If
            Producto.SetFocus
            
        Case 1
            Indice = pantalla.ListIndex
            ClavePrue$ = WIndice.List(Indice)
            spPrueart = "ConsultaPrueart" + "'" + ClavePrue$ + "'"
            Set rstPrueart = db.OpenRecordset(spPrueart, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrueart.RecordCount > 0 Then
                Partida.Text = Mid$(ClavePrue$, 2, 6)
                Producto.Text = rstPrueart!Producto
                fecha.Text = rstPrueart!fecha
                Orden.Text = rstPrueart!Orden
                Valor1.Text = rstPrueart!Valor1
                Valor2.Text = rstPrueart!Valor2
                Valor3.Text = rstPrueart!Valor3
                Valor4.Text = rstPrueart!Valor4
                Valor5.Text = rstPrueart!Valor5
                Valor6.Text = rstPrueart!Valor6
                Valor7.Text = rstPrueart!Valor7
                Valor8.Text = rstPrueart!Valor8
                Valor9.Text = rstPrueart!Valor9
                Valor10.Text = rstPrueart!Valor10
                Valor11.Text = IIf(IsNull(rstPrueart!Valor11), "", rstPrueart!Valor11)
                Valor12.Text = IIf(IsNull(rstPrueart!Valor12), "", rstPrueart!Valor12)
                Valor13.Text = IIf(IsNull(rstPrueart!Valor13), "", rstPrueart!Valor13)
                Valor14.Text = IIf(IsNull(rstPrueart!Valor14), "", rstPrueart!Valor14)
                Valor15.Text = IIf(IsNull(rstPrueart!Valor15), "", rstPrueart!Valor15)
                Valor16.Text = IIf(IsNull(rstPrueart!Valor16), "", rstPrueart!Valor16)
                Valor17.Text = IIf(IsNull(rstPrueart!Valor17), "", rstPrueart!Valor17)
                Valor18.Text = IIf(IsNull(rstPrueart!Valor18), "", rstPrueart!Valor18)
                Valor19.Text = IIf(IsNull(rstPrueart!Valor19), "", rstPrueart!Valor19)
                Valor20.Text = IIf(IsNull(rstPrueart!Valor20), "", rstPrueart!Valor20)
                Valor21.Text = IIf(IsNull(rstPrueart!Valor21), "", rstPrueart!Valor21)
                Valor22.Text = IIf(IsNull(rstPrueart!Valor22), "", rstPrueart!Valor22)
                Valor23.Text = IIf(IsNull(rstPrueart!Valor23), "", rstPrueart!Valor23)
                Valor24.Text = IIf(IsNull(rstPrueart!Valor24), "", rstPrueart!Valor24)
                Valor25.Text = IIf(IsNull(rstPrueart!Valor25), "", rstPrueart!Valor25)
                Valor26.Text = IIf(IsNull(rstPrueart!Valor26), "", rstPrueart!Valor26)
                Valor27.Text = IIf(IsNull(rstPrueart!Valor27), "", rstPrueart!Valor27)
                Valor28.Text = IIf(IsNull(rstPrueart!Valor28), "", rstPrueart!Valor28)
                Valor29.Text = IIf(IsNull(rstPrueart!Valor29), "", rstPrueart!Valor29)
                Valor30.Text = IIf(IsNull(rstPrueart!Valor30), "", rstPrueart!Valor30)
                
                Rem Descriprod.Caption = ""
                Rem Descri1.Caption = ""
                Ensayo.Text = rstPrueart!Ensayo
                Aspecto.Text = rstPrueart!Aspecto
                Observaciones.Text = rstPrueart!Observaciones
                Confecciono.Text = rstPrueart!Confecciono
                Rem Std1.Caption = ""
                Auxi = Left$(rstPrueart!Prueba, 1)
                Lote.Text = Right$(rstPrueart!Prueba, 6)
                Liberada.Text = rstPrueart!Liberada
                Devuelta.Text = rstPrueart!Devuelta
                NroRechazo.Text = rstPrueart!Rechazo
                Nueva.Text = rstPrueart!Nueva
                
                rstPrueart.Close
                
                XEmpresa = WEmpresa
                Select Case Val(WEmpresa)
                    Case 1, 3, 5, 6, 7, 10, 11
                        WEmpresa = "0003"
                        txtOdbc = "Empresa03"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                    Case Else
                        WEmpresa = "0004"
                        txtOdbc = "Empresa04"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                End Select
                
                LlamaImprime = "N"
                
                Sql1 = "Select *"
                Sql2 = " FROM EspecificacionesUnifica"
                Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Producto.Text + "'"
                spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnifica.RecordCount > 0 Then
                    rstEspecificacionesUnifica.Close
                    LlamaImprime = "S"
                End If
                
                Call Conecta_Empresa
                
                If LlamaImprime = "S" Then
                    Call imprime_Click
                End If
                
                spArticulo = "ConsultaArticulo " + "'" + Producto.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Descriprod.Caption = rstArticulo!Descripcion
                    rstArticulo.Close
                End If
                
                spLaudo = "ListaLaudo " + "'" + Lote.Text + "'"
                Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
                If rstLaudo.RecordCount > 0 Then
                    OrigenMercaderiaII.Text = rstLaudo!Origen
                    PartidaProveedorII.Text = rstLaudo!partiori
                    rstLaudo.Close
                End If
                    
                    Else
                    
                Call CmdLimpiar_Click
                
            End If
            Producto.SetFocus
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgPruart.Caption = "Ingreso de Ensayos de Materia Prima :  " + !Nombre
        End If
    End With
    EmpresaActual = WEmpresa
    fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    lblresultado.Caption = "Valor Standard"
    lblresultadoII.Caption = "Valor Standard"
End Sub

Private Sub Cambio_Click()
    WProceso = 0
    Pass.Visible = True
    WClave.Text = ""
    WClave.SetFocus
End Sub

Private Sub Modif_Cancela_Click()
    Modif.Visible = False
End Sub

Private Sub Modif_Orden_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spOrden = "ListaOrden " + "'" + Modif_Orden.Text + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            rstOrden.Close
            Modif_Solicitado.SetFocus
                Else
            m$ = "Orden de Compra Inexistente"
            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
            Modif_Orden.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Modif_Solicitado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Modif_Solicitado.Text <> "" Then
            Modif_Solicitado.Text = UCase(Modif_Solicitado.Text)
            spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                rstArticulo.Close
            
                Llave = "N"
                For WDa% = 1 To 40
                    Auxi3 = Modif_Orden.Text
                    Call Ceros(Auxi3, 6)
                    Auxi1 = WDa%
                    Call Ceros(Auxi1, 2)
                    WClave = Auxi3 + Auxi1
                    spOrden = "ConsultaOrden " + "'" + WClave + "'"
                    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
                    If rstOrden.RecordCount > 0 Then
                        If Modif_Solicitado.Text = rstOrden!Articulo Then
                            Llave = "S"
                        End If
                        rstOrden.Close
                    End If
                Next WDa%
    
                Select Case Llave
                    Case "S"
                        Modif_Recibido.SetFocus
                    Case "N"
                        m$ = "No existe el articulo en la orden de compra especificada"
                        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                        Modif_Solicitado.SetFocus
                    Case Else
                End Select
                    Else
                m$ = "No existe el articulo especificado"
                A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                Modif_Solicitado.SetFocus
            End If
        End If
    End If
End Sub
    
Sub Modif_Recibido_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Modif_Recibido.Text <> "" Then
            Modif_Recibido.Text = UCase(Modif_Recibido.Text)
            spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido, 6) Then
                    m$ = "El articulo recibido no es igual al solicitado"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    Modif_Recibido.SetFocus
                        Else
                    Modif_Orden.SetFocus
                End If
                rstArticulo.Close
                    Else
                If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido, 6) Then
                    m$ = "El articulo recibido no es igual al solicitado"
                    A% = MsgBox(m$, 0, "Ingreso de Pruebas")
                    Modif_Solicitado.SetFocus
                        Else
                    T$ = "Ingreso de Pruebas"
                    m$ = "No existe el articulo especificado, Desea darlo de alta"
                    Respuesta% = MsgBox(m$, 32 + 4, T$)
                    If Respuesta% = 6 Then
                    
                        spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                        If rstArticulo.RecordCount > 0 Then
                        
                            WCodigo = Modif_Recibido.Text
                            WDescripcion = rstArticulo!Descripcion
                            Wunidad = rstArticulo!Unidad
                            WDeposito = rstArticulo!Deposito
                            WInicial = ""
                            WEntradas = ""
                            WSalidas = ""
                            WMinimo = ""
                            WLaboratorio = ""
                            WPedido = ""
                            WEnvase = Str$(rstArticulo!Envase)
                            WCosto1 = Str$(rstArticulo!Costo1)
                            WCosto2 = Str$(rstArticulo!Costo2)
                            WRs = rstArticulo!Rs
                            WFlete = Str$(rstArticulo!Flete)
                            WMoneda = rstArticulo!Moneda
                            WControla = Str$(IIf(IsNull(rstArticulo!Controla), "0", rstArticulo!Controla))
                            WDensidad = IIf(IsNull(rstArticulo!Densidad), "", rstArticulo!Densidad)
                            WProveedor = IIf(IsNull(rstArticulo!Proveedor), "", rstArticulo!Proveedor)
                            WDate = Date$
                            WFecha = IIf(IsNull(rstArticulo!fecha), "", rstArticulo!fecha)
                            WOrden = IIf(IsNull(rstArticulo!Orden), "", Str$(rstArticulo!Orden))
                            WDife = ""
                            WCosto3 = ""
                            
                            rstArticulo.Close
                            
                            XParam = "'" + WCodigo + "','" _
                                + WDescripcion + "','" _
                                + WCosto1 + "','" _
                                + WCosto2 + "','" _
                                + WInicial + "','" _
                                + WEntradas + "','" _
                                + WSalidas + "','" _
                                + WMinimo + "','" _
                                + WLaboratorio + "','" _
                                + Wunidad + "','" _
                                + WPedido + "','" _
                                + WDeposito + "','" _
                                + WEnvase + "','" _
                                + WRs + "','" _
                                + WFecha + "','" _
                                + WOrden + "','" _
                                + WDife + "','" _
                                + WProveedor + "','" _
                                + WDate + "','" _
                                + WFlete + "','" _
                                + WMoneda + "','" _
                                + WControla + "','" _
                                + WDensidad + "','" _
                                + WCosto3 + "'"
                         
                            Set rstArticulo = db.OpenRecordset("AltaArticulo " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub Modif_Confirma_Click()

    Trabajo = "S"
    
    spOrden = "ListaOrden " + "'" + Modif_Orden.Text + "'"
    Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
    If rstOrden.RecordCount > 0 Then
        rstOrden.Close
            Else
        Trabajo = "N"
        m$ = "Orden de Compra Inexistente"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Trabajo = "N"
        m$ = "No existe el articulo especificado en Articulo Pedido"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        rstArticulo.Close
            Else
        Trabajo = "N"
        m$ = "No existe el articulo especificado en Articulo Recibido"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    Llave = "N"
    For WDa% = 1 To 40
        Auxi3 = Modif_Orden.Text
        Call Ceros(Auxi3, 6)
        Auxi1 = WDa%
        Call Ceros(Auxi1, 2)
        WClave = Auxi3 + Auxi1
        spOrden = "ConsultaOrden " + "'" + WClave + "'"
        Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
        If rstOrden.RecordCount > 0 Then
            If Modif_Solicitado.Text = rstOrden!Articulo Then
                Llave = "S"
            End If
            rstOrden.Close
        End If
    Next WDa%
    
    Select Case Llave
        Case "N"
            Trabajo = "N"
            m$ = "No existe el articulo en la orden de compra especificada"
            A% = MsgBox(m$, 0, "Ingreso de Pruebas")
        Case Else
    End Select
    
    If Left$(Modif_Solicitado.Text, 6) <> Left$(Modif_Recibido.Text, 6) Then
        Trabajo = "N"
        m$ = "El articulo recibido no es igual al solicitado"
        A% = MsgBox(m$, 0, "Ingreso de Pruebas")
    End If

    If Trabajo = "S" Then
    
        Modif.Visible = False
        
        For WDa% = 1 To 40
            Auxi3 = Modif_Orden.Text
            Call Ceros(Auxi3, 6)
            Auxi1 = WDa%
            Call Ceros(Auxi1, 2)
            WClave = Auxi3 + Auxi1
            spOrden = "ConsultaOrden " + "'" + WClave + "'"
            Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)
            If rstOrden.RecordCount > 0 Then
                If Modif_Solicitado.Text = rstOrden!Articulo Then
                
                    WCantidad = rstOrden!Cantidad
                    WClave = rstOrden!Clave
                    WArticulo = Modif_Recibido.Text
                    WDate = Date$
                    rstOrden.Close
                    XParam = "'" + WClave + "','" _
                                + WArticulo + "','" _
                                + WDate + "'"
                    Set rstOrden = db.OpenRecordset("ModificaOrdenLaboratorio " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    
                    spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = Modif_Solicitado.Text
                        WPedido = Str$(rstArticulo!Pedido - WCantidad)
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WPedido + "','" _
                                + WDate + "'"
                        spArticulo = "ModificaArticuloOrdenLaboratorio " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                    spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
                    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    If rstArticulo.RecordCount > 0 Then
                        WCodigo = Modif_Recibido.Text
                        WPedido = Str$(rstArticulo!Pedido + WCantidad)
                        WDate = Date$
                        rstArticulo.Close
                        XParam = "'" + WCodigo + "','" _
                                + WPedido + "','" _
                                + WDate + "'"
                        spArticulo = "ModificaArticuloOrdenLaboratorio " + XParam
                        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                    End If
                    
                        Else
                        
                    rstOrden.Close
                    
                End If
            End If
        Next WDa%
        
        XParam = "'" + Modif_Orden.Text + "','" _
                    + Modif_Solicitado.Text + "'"
        spInforme = "ListaInformeOrdenArticulo " + XParam
        Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
        If rstInforme.RecordCount > 0 Then
            WCantidad = rstInforme!Cantidad
            WResta = rstInforme!Resta
            WClave = rstInforme!Clave
            WArticulo = Modif_Recibido.Text
            WDate = Date$
            rstInforme.Close
            XParam = "'" + WClave + "','" _
                        + WArticulo + "','" _
                        + WDate + "'"
            Set rstInforme = db.OpenRecordset("ModificaInformeLaboratorio " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            
            spArticulo = "ConsultaArticulo " + "'" + Modif_Solicitado.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = Modif_Solicitado.Text
                WPedido = Str$(rstArticulo!Pedido + WResta)
                WLaboratorio = Str$(rstArticulo!Laboratorio - WCantidad)
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WLaboratorio + "','" _
                            + WDate + "'"
                spArticulo = "ModificaArticuloInformeLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Modif_Recibido.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                WCodigo = Modif_Recibido.Text
                WPedido = Str$(rstArticulo!Pedido - WResta)
                WLaboratorio = Str$(rstArticulo!Laboratorio + WCantidad)
                WDate = Date$
                rstArticulo.Close
                XParam = "'" + WCodigo + "','" _
                            + WPedido + "','" _
                            + WLaboratorio + "','" _
                            + WDate + "'"
                spArticulo = "ModificaArticuloInformeLaboratorio " + XParam
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
        End If
        
        Producto.Text = Modif_Recibido.Text
        Producto.SetFocus
        
    End If
    
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case WProceso
            Case 0
                If WClave.Text = "MATIZ" Then
                    Pass.Visible = False
                    Modif_Orden.Text = ""
                    Modif_Solicitado.Text = "  -   -   "
                    Modif_Recibido.Text = "  -   -   "
                    Modif.Visible = True
                    Modif_Orden.SetFocus
                End If
            Case Else
                If WClave.Text = "SEGURO" Then
                    Pass.Visible = False
                    Call ModificaPrueba
                End If
        End Select
    End If
End Sub

Private Sub WCancela_Click()
    Pass.Visible = False
End Sub

Private Sub Calcula_SaldoOrden()

    WRecibida = 0
    WLaudada = 0

    spInforme = "ListaInformeOrden " + "'" + Orden.Text + "'"
    Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
    If rstInforme.RecordCount > 0 Then
        With rstInforme
            .MoveFirst
            Do
                If Producto.Text = rstInforme!Articulo Then
                    WRecibida = WRecibida + rstInforme!Cantidad
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstInforme.Close
    End If
    
    spLaudo = "ListaLaudoOrden " + "'" + Orden.Text + "'"
    Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
    If rstLaudo.RecordCount > 0 Then
        With rstLaudo
            .MoveFirst
            Do
                If Producto.Text = rstLaudo!Articulo Then
                    WCantidad1 = rstLaudo!Liberada
                    WCantidad2 = rstLaudo!Devuelta
                    WLaudada = WLaudada + WCantidad1 + WCantidad2
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End With
        rstLaudo.Close
    End If
    
    SaldoOrden = WRecibida - WLaudada

End Sub


Public Sub SendEmail()

    Dim objOutlook As Object
    Dim objMailItem

    Dim NumOfPath As Integer, i As Integer
    Dim AtachPath As String

    On Error GoTo 10

    NumOfPath = 0
    AllPath = ""
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMailItem = objOutlook.CreateItem(olMailItem)
    
    With objMailItem
        .To = EmailAddress
        .cc = CopiaAddress
        .Subject = MSubject
        .Body = MBody
        Rem .Attachments.Add MAttach
        Rem If MAttachI <> "" Then
        Rem     .Attachments.Add MAttachI
        Rem End If
        Rem If MAttachII <> "" Then
        Rem     .Attachments.Add MAttachII
        Rem End If
        Rem If MAttachIII > "" Then
        Rem     .Attachments.Add MAttachIII
        Rem End If
        Rem If MAttachIV <> "" Then
        Rem     .Attachments.Add MAttachIV
        Rem End If
        Rem If MAttachV <> "" Then
        Rem     .Attachments.Add MAttachV
        Rem End If
        Rem If MAttachVI <> "" Then
        Rem     .Attachments.Add MAttachVI
        Rem End If
        Rem If MAttachVII <> "" Then
        Rem     .Attachments.Add MAttachVII
        Rem End If
        Rem If MAttachVIII <> "" Then
        Rem     .Attachments.Add MAttachVIII
        Rem End If
        .Send
    End With

    Set objMailItem = Nothing
    Set objOutlook = Nothing
            
    Exit Sub

exit10:
    Exit Sub

10:
    If Err.Number = 429 Then
        MsgBox "Error on connecting with Outlook"
            Else
        MsgBox "error Description is  " & Err.Description & " err number is " & Err.Number
    End If
    Set objMailItem = Nothing
    Set objOutlook = Nothing
    AllPath = ""

    Resume exit10

End Sub

Sub Control_Controla(KeyAscii As Integer)
    If KeyAscii <> Asc("C") And KeyAscii <> Asc("N") Then
        KeyAscii = 0   'discard it
    End If
End Sub

