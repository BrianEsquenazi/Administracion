VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgPrecio 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ingreso de Precios por Cliente"
   ClientHeight    =   8160
   ClientLeft      =   915
   ClientTop       =   435
   ClientWidth     =   10170
   LinkTopic       =   "Form2"
   ScaleHeight     =   8160
   ScaleWidth      =   10170
   Begin VB.Frame Clave1 
      Caption         =   "  Ingreso de Clave de Seguridad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   5280
      TabIndex        =   97
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton Cancelagraba 
         Caption         =   "Cancela Grabacion"
         Height          =   255
         Left            =   960
         TabIndex        =   99
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   98
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Ingrese su Password"
         Height          =   255
         Left            =   1080
         TabIndex        =   100
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame PantaPt 
      Height          =   7335
      Left            =   120
      TabIndex        =   50
      Top             =   480
      Visible         =   0   'False
      Width           =   5295
      Begin VB.Frame PantaClaveNombre 
         Caption         =   "  Ingreso de Clave de Seguridad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   1200
         TabIndex        =   122
         Top             =   2400
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox WClaveNombre 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   960
            PasswordChar    =   "*"
            TabIndex        =   124
            Top             =   600
            Width           =   1695
         End
         Begin VB.CommandButton CancelaGrabaNombre 
            Caption         =   "Cancela Grabacion"
            Height          =   255
            Left            =   960
            TabIndex        =   123
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Ingrese su Password"
            Height          =   255
            Left            =   1080
            TabIndex        =   125
            Top             =   360
            Width           =   1815
         End
      End
      Begin MSFlexGridLib.MSFlexGrid Muestra4 
         Height          =   1695
         Left            =   1920
         TabIndex        =   108
         Top             =   2520
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   2990
         _Version        =   327680
         Rows            =   100
         BackColor       =   16777088
      End
      Begin VB.Frame PantaModifNombre 
         Height          =   2655
         Left            =   1560
         TabIndex        =   113
         Top             =   1920
         Visible         =   0   'False
         Width           =   5175
         Begin VB.CommandButton CancelaModif 
            Caption         =   "Cancela"
            Height          =   495
            Left            =   2520
            TabIndex        =   109
            Top             =   1920
            Width           =   1215
         End
         Begin VB.CommandButton ConfirmaModif 
            Caption         =   "Confirma"
            Height          =   495
            Left            =   960
            TabIndex        =   116
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox NombreI 
            Height          =   315
            Left            =   240
            TabIndex        =   115
            Top             =   600
            Width           =   4575
         End
         Begin VB.ComboBox NombreII 
            Height          =   315
            Left            =   240
            TabIndex        =   114
            Top             =   1440
            Width           =   4575
         End
         Begin VB.Label Label21 
            Caption         =   "Nueva Descripcion"
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
            TabIndex        =   119
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label Label20 
            Caption         =   "Descripcion a Reemplazar"
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
            TabIndex        =   118
            Top             =   240
            Width           =   3615
         End
      End
      Begin VB.Frame MenuNOmbre 
         BackColor       =   &H00FFFFC0&
         Height          =   3375
         Left            =   480
         TabIndex        =   126
         Top             =   3120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.CommandButton CancelaMenu 
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
            Left            =   360
            TabIndex        =   130
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CommandButton BotomNuevo 
            Caption         =   "Nuevo Nombre"
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
            Left            =   360
            TabIndex        =   129
            Top             =   360
            Width           =   1215
         End
         Begin VB.CommandButton Reproceso 
            Caption         =   "Reproceso"
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
            Left            =   360
            TabIndex        =   128
            Top             =   1800
            Width           =   1215
         End
         Begin VB.CommandButton BotonModif 
            Caption         =   "Modifica Nombre"
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
            Left            =   360
            TabIndex        =   127
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.CommandButton Mantenimiento 
         Caption         =   "Mantenimiento de Nombres"
         Height          =   615
         Left            =   120
         TabIndex        =   121
         Top             =   3000
         Width           =   1335
      End
      Begin VB.Frame PantaAltaNombre 
         Height          =   1695
         Left            =   240
         TabIndex        =   110
         Top             =   1560
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton CancelaAlta 
            Caption         =   "Cancela"
            Height          =   495
            Left            =   2640
            TabIndex        =   117
            Top             =   960
            Width           =   1215
         End
         Begin VB.CommandButton ConfirmaAlta 
            Caption         =   "Confirma"
            Height          =   495
            Left            =   960
            TabIndex        =   112
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox NombreAlta 
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
            MaxLength       =   50
            TabIndex        =   111
            Top             =   600
            Width           =   4575
         End
         Begin VB.Label Label22 
            Caption         =   "Nueva Descripcion"
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
            TabIndex        =   120
            Top             =   240
            Width           =   1935
         End
      End
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
         Left            =   3600
         TabIndex        =   105
         Top             =   960
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Height          =   3015
         Left            =   3240
         TabIndex        =   72
         Top             =   4200
         Visible         =   0   'False
         Width           =   5295
         Begin VB.TextBox ListaVendedor 
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
            TabIndex        =   101
            Text            =   " "
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox DesdeCliente 
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
            MaxLength       =   6
            TabIndex        =   78
            Text            =   " "
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox HastaCliente 
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
            MaxLength       =   6
            TabIndex        =   77
            Text            =   " "
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton Cancela 
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
            Height          =   375
            Left            =   3720
            TabIndex        =   76
            Top             =   600
            Width           =   975
         End
         Begin VB.CommandButton Acepta 
            Caption         =   "Acepta"
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
            Left            =   3720
            TabIndex        =   75
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton Impresora 
            Caption         =   "Impresora"
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
            Left            =   2760
            TabIndex        =   74
            Top             =   2450
            Width           =   1215
         End
         Begin VB.OptionButton Panta 
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
            Height          =   255
            Left            =   1320
            TabIndex        =   73
            Top             =   2520
            Width           =   1095
         End
         Begin MSMask.MaskEdBox HastaTerminado 
            Height          =   375
            Left            =   1800
            TabIndex        =   79
            Top             =   1560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   327680
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "AA-#####-###"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox DesdeTerminado 
            Height          =   375
            Left            =   1800
            TabIndex        =   80
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   327680
            MaxLength       =   12
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "AA-#####-###"
            PromptChar      =   " "
         End
         Begin VB.Label Label10 
            Caption         =   "Vendedor"
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
            TabIndex        =   102
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Desde Cliente"
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
            TabIndex        =   84
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label5 
            Caption         =   "Desde Producto"
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
            TabIndex        =   83
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta Producto"
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
            TabIndex        =   82
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "Hasta Cliente"
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
            TabIndex        =   81
            Top             =   720
            Width           =   1455
         End
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
         Height          =   1980
         Left            =   2760
         TabIndex        =   71
         Top             =   4560
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.TextBox Descripcion 
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   70
         Text            =   " "
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   65
         Top             =   5400
         Width           =   2055
         Begin VB.CommandButton Primer 
            Caption         =   "Primer Reg."
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
            TabIndex        =   69
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton Ultimo 
            Caption         =   "Ultimo Reg."
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
            TabIndex        =   68
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton Siguiente 
            Caption         =   "Reg. Siguiente"
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
            TabIndex        =   67
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton Anterior 
            Caption         =   "Reg. Anterior"
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
            TabIndex        =   66
            Top             =   960
            Width           =   1815
         End
      End
      Begin VB.ListBox Pantalla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         ItemData        =   "precios.frx":0000
         Left            =   2400
         List            =   "precios.frx":0007
         TabIndex        =   64
         Top             =   4560
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
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
         Left            =   120
         TabIndex        =   63
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   62
         Top             =   4920
         Width           =   975
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
         Height          =   300
         Left            =   1200
         TabIndex        =   61
         Top             =   4560
         Width           =   975
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
         Height          =   300
         Left            =   120
         TabIndex        =   60
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton lista 
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
         Height          =   300
         Left            =   1200
         TabIndex        =   59
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton CmdLimpiar 
         Caption         =   "Limpiar"
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
         Left            =   1200
         TabIndex        =   58
         Top             =   4920
         Width           =   975
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
         Left            =   2400
         TabIndex        =   57
         Top             =   4200
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.TextBox Cliente 
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
         Left            =   2040
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Precio 
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
         Left            =   2040
         TabIndex        =   56
         Text            =   " "
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Pago 
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
         MaxLength       =   4
         TabIndex        =   55
         Top             =   1680
         Width           =   975
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
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   2880
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
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   2880
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
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2880
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2880
         Width           =   375
      End
      Begin MSMask.MaskEdBox Terminado 
         Height          =   285
         Left            =   2040
         TabIndex        =   85
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "AA-#####-###"
         PromptChar      =   " "
      End
      Begin MSFlexGridLib.MSFlexGrid WVector1 
         Height          =   1935
         Left            =   1560
         TabIndex        =   86
         Top             =   2160
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label Label14 
         Caption         =   "Descripcion"
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
         TabIndex        =   107
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Descripcion"
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
         TabIndex        =   96
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cliente"
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
         Left            =   120
         TabIndex        =   95
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H0080FFFF&
         Caption         =   "Producto Terminado"
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
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   94
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Precio"
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
         TabIndex        =   93
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label DesCliente 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3600
         TabIndex        =   92
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label DesTerminado 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3600
         TabIndex        =   91
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha Mod."
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
         TabIndex        =   90
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Fecha 
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
         Left            =   2040
         TabIndex        =   89
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Condicion de Pago"
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
         Left            =   3720
         TabIndex        =   88
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label DesPago 
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
         Left            =   6600
         TabIndex        =   87
         Top             =   1680
         Width           =   2895
      End
   End
   Begin VB.Frame PantaDy 
      Height          =   7575
      Left            =   5640
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox Estado1 
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
         Left            =   3720
         TabIndex        =   106
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox WTitulo1 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Frame Frame5 
         Height          =   3015
         Left            =   3480
         TabIndex        =   22
         Top             =   3840
         Visible         =   0   'False
         Width           =   5295
         Begin VB.TextBox ListaVendedor1 
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
            TabIndex        =   103
            Text            =   " "
            Top             =   2040
            Width           =   975
         End
         Begin VB.OptionButton Panta1 
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
            Height          =   255
            Left            =   1320
            TabIndex        =   28
            Top             =   2520
            Width           =   1095
         End
         Begin VB.OptionButton Impresora1 
            Caption         =   "Impresora"
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
            Left            =   2880
            TabIndex        =   27
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton Acepta1 
            Caption         =   "Acepta"
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
            Left            =   3720
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.CommandButton Cancela1 
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
            Height          =   375
            Left            =   3720
            TabIndex        =   25
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox HastaCliente1 
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
            MaxLength       =   6
            TabIndex        =   24
            Text            =   " "
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox DesdeCliente1 
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
            MaxLength       =   6
            TabIndex        =   23
            Text            =   " "
            Top             =   360
            Width           =   975
         End
         Begin MSMask.MaskEdBox HastaArticulo 
            Height          =   375
            Left            =   1800
            TabIndex        =   29
            Top             =   1560
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin MSMask.MaskEdBox DesdeArticulo 
            Height          =   375
            Left            =   1800
            TabIndex        =   30
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
         Begin VB.Label Label11 
            Caption         =   "Vendedor"
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
            TabIndex        =   104
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label Label19 
            Caption         =   "Hasta Cliente"
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
            TabIndex        =   34
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "Hasta M.P."
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
            TabIndex        =   33
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label17 
            Caption         =   "Desde M.P."
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
            TabIndex        =   32
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label Label12 
            Caption         =   "Desde Cliente"
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
            TabIndex        =   31
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.ListBox Opcion1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   2880
         TabIndex        =   21
         Top             =   4200
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.TextBox Precio1 
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
         Left            =   2160
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Pago1 
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
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox Cliente1 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   18
         Text            =   " "
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Ayuda1 
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
         TabIndex        =   17
         Top             =   3840
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.CommandButton CmdLimpiar1 
         Caption         =   "Limpiar"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton Lista1 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton Consulta1 
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
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton CmdClose1 
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
         Height          =   300
         Left            =   1320
         TabIndex        =   13
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton CmdDelete1 
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
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
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   975
      End
      Begin VB.CommandButton cmdadd1 
         Caption         =   "&Agregar"
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
         Left            =   240
         TabIndex        =   11
         Top             =   3840
         Width           =   975
      End
      Begin VB.ListBox Pantalla1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2700
         ItemData        =   "precios.frx":0015
         Left            =   2520
         List            =   "precios.frx":001C
         TabIndex        =   10
         Top             =   4200
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   240
         TabIndex        =   5
         Top             =   5040
         Width           =   2055
         Begin VB.CommandButton Anterior1 
            Caption         =   "Reg. Anterior"
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
            Top             =   960
            Width           =   1815
         End
         Begin VB.CommandButton Siguiente1 
            Caption         =   "Reg. Siguiente"
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
            TabIndex        =   8
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton Ultimo1 
            Caption         =   "Ultimo Reg."
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
            TabIndex        =   7
            Top             =   480
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton Primer1 
            Caption         =   "Primer Reg."
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
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1815
         End
      End
      Begin MSMask.MaskEdBox Articulo 
         Height          =   285
         Left            =   2160
         TabIndex        =   39
         Top             =   720
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
      Begin MSFlexGridLib.MSFlexGrid WVector2 
         Height          =   1935
         Left            =   1560
         TabIndex        =   40
         Top             =   1800
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3413
         _Version        =   327680
         BackColor       =   16777152
      End
      Begin VB.Label Despago1 
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
         Left            =   6720
         TabIndex        =   49
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label15 
         Caption         =   "Condicion de Pago"
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
         Left            =   3840
         TabIndex        =   48
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Fecha1 
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
         Left            =   2160
         TabIndex        =   47
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Fecha Mod."
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
         TabIndex        =   46
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label DesArticulo 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3720
         TabIndex        =   45
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label DesCliente1 
         BackColor       =   &H00FFFF00&
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
         Height          =   285
         Left            =   3720
         TabIndex        =   44
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label9 
         Caption         =   "Precio"
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
         TabIndex        =   43
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H0080FFFF&
         Caption         =   "Materia Prima"
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cliente"
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
         Index           =   2
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton BotonDy 
      Caption         =   "Reventa"
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
      Left            =   5640
      MaskColor       =   &H00808080&
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton BotonPt 
      Caption         =   "Producto Terminado"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   9360
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Wprecios.rpt"
      WindowTitle     =   "Listado de Precios por Cliente"
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8640
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "PrgPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPrecios As Recordset
Dim spPrecios As String
Dim rstPreciosMp As Recordset
Dim spPreciosMp As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstPago As Recordset
Dim spPago As String
Dim rstPreciosII As Recordset
Dim spPreciosII As String

Dim XParam As String

Private WFecha1 As String
Private WFactura1 As String
Private WPrecio1 As String
Private WCantidad1 As String
Private WFecha2 As String
Private WFactura2 As String
Private WPrecio2 As String
Private WCantidad2 As String
Private WFecha3 As String
Private WFactura3 As String
Private WPrecio3 As String
Private WCantidad3 As String
Private WFecha4 As String
Private WFactura4 As String
Private WPrecio4 As String
Private WCantidad4 As String
Private WFecha5 As String
Private WFactura5 As String
Private WPrecio5 As String
Private WCantidad5 As String

Private Dada As String
Dim WGraba As String
Dim WGrabaNombre As String
Dim WGrabaProceso As Integer
Dim ZEstado As Integer
Dim ZImpreVto As Integer
Dim ZVida As Integer

Dim ZZVector(10000, 3) As String

Private Sub Acepta_Click()

    DesdeTerminado.Text = UCase(DesdeTerminado.Text)
    HastaTerminado.Text = UCase(HastaTerminado.Text)
    DesdeCliente.Text = UCase(DesdeCliente.Text)
    Hastacliente.Text = UCase(Hastacliente.Text)
    
    Listado.WindowTitle = "Listado de Precios de Productos Terminados por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Impresora.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    If Val(ListaVendedor.Text) <> 0 Then
        WDesdeVendedor = ListaVendedor.Text
        WHastaVendedor = ListaVendedor.Text
            Else
        WDesdeVendedor = "0"
        WHastaVendedor = "9999"
    End If
    
    Uno = "{Precios.Terminado} in " + Chr$(34) + DesdeTerminado.Text + Chr$(34) + " to " + Chr$(34) + HastaTerminado.Text + Chr$(34)
    Dos = " and " + "{Precios.Cliente} in " + Chr$(34) + DesdeCliente.Text + Chr$(34) + " to " + Chr$(34) + Hastacliente.Text + Chr$(34)
    Tres = " and " + "{Cliente.Vendedor} in " + WDesdeVendedor + " to " + WHastaVendedor
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT Precios.Cliente, Precios.Terminado, Precios.Precio, Precios.Descripcion, " _
                    + "Cliente.Razon, Cliente.Vendedor " _
                    + "From " _
                    + DSQ + ".dbo.Precios Precios, " _
                    + DSQ + ".dbo.Cliente Cliente " _
                    + "Where " _
                    + "Precios.Cliente = Cliente.Cliente AND " _
                    + "Precios.Cliente >= '" + DesdeCliente.Text + "' AND " _
                    + "Precios.Cliente <= '" + Hastacliente.Text + "' AND " _
                    + "Precios.Terminado >= '" + DesdeTerminado.Text + "' AND " _
                    + "Precios.Terminado <= '" + HastaTerminado.Text + "'"
    
    Listado.DataFiles(0) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.ReportFileName = "WPrecios.rpt"
    
    Cliente.SetFocus
    Listado.Action = 1
    Frame2.Visible = False
End Sub


Private Sub BotonDy_Click()
    PantaPt.Visible = False
    PantaDy.Visible = True
    PantaDy.Height = 7575
    PantaDy.Left = 120
    PantaDy.Top = 360
    PantaDy.Width = 9855
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Cliente1.Text = ""
    DesCliente1.Caption = ""
    Precio1.Text = ""
    Fecha1.Caption = ""
    Call Limpia_Vector1
    Cliente1.SetFocus
End Sub

Private Sub BotonModif_Click()
    
    MenuNOmbre.Visible = False
    
    NombreI.Clear
    NombreII.Clear
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosII"
    ZSql = ZSql + " Where PreciosII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by Clave"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosII.RecordCount > 0 Then
    
        With rstPreciosII
            .MoveFirst
            Do
                If .EOF = False Then
    
                    NombreI.AddItem rstPreciosII!Nombre
                    NombreII.AddItem rstPreciosII!Nombre
                    
                    .MoveNext
                        Else
                    Rem BY NAN
                     NombreI.ListIndex = 0
                     NombreII.ListIndex = 0
                    Exit Do
                
              
                End If
            Loop
        End With
        rstPreciosII.Close
    End If
    
    NombreI.ListIndex = -1
   NombreII.ListIndex = -1
    
    PantaModifNombre.Visible = True
    
End Sub

Private Sub CancelaAlta_Click()
    PantaAltaNombre.Visible = False
End Sub

Private Sub ConfirmaAlta_Click()

    ZRenglon = 1

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosII"
    ZSql = ZSql + " Where PreciosII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by PreciosII.Clave"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosII.RecordCount > 0 Then
        With rstPreciosII
            .MoveLast
            ZRenglon = rstPreciosII!Renglon + 1
        End With
        rstPreciosII.Close
    End If


    ZZTerminado = Terminado.Text
    ZZDescripcion = NombreAlta.Text
    ZZRenglon = Str$(ZRenglon)
    
    Auxi = ZZRenglon
    Call Ceros(Auxi, 2)
    ZZClave = ZZTerminado + Auxi
    
    ZSql = ""
    ZSql = ZSql & "INSERT INTO PreciosII ("
    ZSql = ZSql & "Clave ,"
    ZSql = ZSql & "Terminado ,"
    ZSql = ZSql & "Renglon ,"
    ZSql = ZSql & "Nombre )"
    ZSql = ZSql & "Values ("
    ZSql = ZSql & "'" + ZZClave + "',"
    ZSql = ZSql & "'" + ZZTerminado + "',"
    ZSql = ZSql & "'" + ZZRenglon + "',"
    ZSql = ZSql & "'" + ZZDescripcion + "')"

    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)

    PantaAltaNombre.Visible = False
    
End Sub

Private Sub CancelaMenu_Click()
    MenuNOmbre.Visible = False
End Sub

Private Sub BotomNuevo_Click()
    NombreAlta.Text = ""
    MenuNOmbre.Visible = False
    PantaAltaNombre.Visible = True
    NombreAlta.SetFocus
End Sub


Private Sub CancelaModif_Click()
    PantaModifNombre.Visible = False
End Sub

Private Sub ConfirmaModif_Click()

    Dim ZZCambia(1000) As String
    
    Erase ZZCambia
    ZZLugar = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Where Precios.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by Clave"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
    
                    If Trim(UCase(rstPrecios!Descripcion)) = Trim(UCase(NombreI.Text)) Then
                        ZZLugar = ZZLugar + 1
                        ZZCambia(ZZLugar) = rstPrecios!Clave
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
    End If
    
    For ZZCiclo = 1 To ZZLugar
    
        ZZClave = ZZCambia(ZZCiclo)
    
        ZSql = ""
        ZSql = ZSql & "UPDATE Precios SET "
        ZSql = ZSql & "Descripcion = " + "'" + NombreII.Text + "'"
        ZSql = ZSql & " Where Clave = " + "'" + ZZClave + "'"
        spPrecios = ZSql
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)

    Next ZZCiclo

    ZSql = ""
    ZSql = ZSql + "DELETE PreciosII"
    ZSql = ZSql + " Where Nombre = " + "'" + NombreI.Text + "'"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)

    PantaModifNombre.Visible = False
End Sub

Private Sub BotonPt_Click()
    PantaDy.Visible = False
    PantaPt.Height = 7575
    PantaPt.Left = 120
    PantaPt.Top = 360
    PantaPt.Width = 9855
    PantaPt.Visible = True
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Precio.Text = ""
    Descripcion.Text = ""
    Fecha.Caption = ""
    Call Limpia_Vector
    Cliente.SetFocus
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Sub Imprime_Descripcion()

    Rem lee Cliente

    WCliente = Cliente.Text
    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente.Caption = ""
    End If
    
    Rem lee Terminado
    
    WTerminado = Terminado.Text
    spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    If rstTerminado.RecordCount > 0 Then
        DesTerminado.Caption = rstTerminado!Descripcion
        rstTerminado.Close
            Else
        DesTerminado.Caption = ""
    End If
    
    WPago = Pago.Text
    spPago = "ConsultaPago " + "'" + Pago.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        DesPago.Caption = rstPago!Nombre
        rstPago.Close
            Else
        DesPago.Caption = ""
    End If
    
End Sub

Sub Verifica_datos()
    If Val(Precio.Text) = 0 Then
        Precio.Text = "0"
    End If
    If Val(Pago.Text) = 0 Then
        Pago.Text = "0"
    End If
End Sub

Sub Format_datos()
    Precio.Text = Pusing("###,###.##", Precio.Text)
End Sub

Sub Imprime_Datos()

    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        Cliente.Text = rstPrecios!Cliente
        Terminado.Text = rstPrecios!Terminado
        Precio.Text = rstPrecios!Precio
        
        If Val(Wempresa) <> 1 Then
            Descripcion.Text = rstPrecios!Descripcion
                Else
            XCodigo = Val(Mid$(Terminado.Text, 4, 5))
            If XCodigo >= 25000 And XCodigo <= 25999 Then
                Descripcion.Text = IIf(IsNull(rstPrecios!descripcionfarma), "", rstPrecios!descripcionfarma)
                    Else
                Descripcion.Text = rstPrecios!Descripcion
            End If
        End If
        
        Fecha.Caption = IIf(IsNull(rstPrecios!Fecha), "", rstPrecios!Fecha)
        Pago.Text = IIf(IsNull(rstPrecios!Pago), "0", rstPrecios!Pago)
        ZEstado = IIf(IsNull(rstPrecios!Estado), "0", rstPrecios!Estado)
        Estado.ListIndex = ZEstado
        Call Format_datos
            
        'columna 1
        
        Call Limpia_Vector
                    
        WVector1.Row = 1
    
        If rstPrecios!Cantidad1 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha1
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura1
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio1)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad1)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 2
    
        WVector1.Row = 2
            
        If rstPrecios!Cantidad2 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha2
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura2
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio2)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad2)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 3
    
        WVector1.Row = 3
            
        If rstPrecios!Cantidad3 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha3
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura3
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio3)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad3)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 4
    
        WVector1.Row = 4
            
        If rstPrecios!Cantidad4 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha4
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura4
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio4)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad4)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 5
    
        WVector1.Row = 5
        
        If rstPrecios!Cantidad5 <> 0 Then
            WVector1.Col = 1
            WVector1.Text = rstPrecios!Fecha5
            WVector1.Col = 2
            WVector1.Text = rstPrecios!Factura5
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", rstPrecios!Precio5)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", rstPrecios!Cantidad5)
                Else
            WVector1.Col = 1
            WVector1.Text = ""
            WVector1.Col = 2
            WVector1.Text = ""
            WVector1.Col = 3
            WVector1.Text = Pusing("###,###.##", Dada)
            WVector1.Col = 4
            WVector1.Text = Pusing("###,###.##", Dada)
        End If
        rstPrecios.Close

    End If
    Call Imprime_Descripcion
    
End Sub


Private Sub cmdAdd_Click()

    If WGraba <> "S" Then
    
        WGrabaProceso = 0
        Call Ingresa_clave
        
            Else
            
        WGraba = ""
        
        If Left$(Terminado.Text, 2) <> "PT" And Left$(Terminado.Text, 2) <> "PE" And Left$(Terminado.Text, 2) <> "YQ" And Left$(Terminado.Text, 2) <> "YF" And Left$(Terminado.Text, 2) <> "YP" And Left$(Terminado.Text, 2) <> "YH" Then
            m$ = "Producto Terminado no esta autorizado para la venta"
            a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
            Exit Sub
        End If

        spTerminado = "ConsultaTerminado " + "'" + UCase(Terminado.Text) + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            rstTerminado.Close
                Else
            m$ = "Producto Terminado inexistente"
            a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
            Exit Sub
        End If

        If Cliente.Text <> "" And Terminado.Text <> "" Then
    
            Cliente.Text = UCase(Cliente.Text)
            Terminado.Text = UCase(Terminado.Text)
    
            WCliente = Cliente.Text
            WTerminado = Terminado.Text
            WClave = Cliente.Text + Terminado.Text
            
            ZImpreVto = 0
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                rstCliente.Close
            End If
            
            If ZImpreVto = 1 Then
            
                ZVida = 0
                spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    ZVida = IIf(IsNull(rstTerminado!Vida), "0", rstTerminado!Vida)
                    rstTerminado.Close
                End If
                
                If ZVida = 0 Then
                    m$ = "Atencion: El producto terminado no posee vida util y el cliente lo exige"
                    a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
                End If
            
            End If
        
        
            If Val(Wempresa) = 1 Then
                XCodigo = Val(Mid$(Terminado.Text, 4, 5))
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    ZDescripcion = Trim(DesTerminado.Caption) + " - " + Trim(Descripcion.Text)
                    If Len(ZDescripcion) > 50 Then
                        m$ = "Al ser el producto de farma la descricion + la descripcin adicional superan los 50 caracteres"
                        a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
                        Exit Sub
                    End If
                End If
            End If

        
        
        
        
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                rstPrecios.Close
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
        
            Call Verifica_datos
        
            WVector1.Row = 1
            WVector1.Col = 4
            Auxi = Val(WVector1.Text)
    
            If Auxi <> 0 Then
                WVector1.Col = 1
                WFecha1 = WVector1.Text
                WVector1.Col = 2
                WFactura1 = WVector1.Text
                WVector1.Col = 3
                WPrecio1 = WVector1.Text
                WVector1.Col = 4
                WCantidad1 = WVector1.Text
                    Else
                WFecha1 = ""
                WFactura1 = ""
                WPrecio1 = ""
                WCantidad1 = ""
            End If
            
            WVector1.Row = 2
            WVector1.Col = 4
            Auxi = Val(WVector1.Text)
    
            If Auxi <> 0 Then
                WVector1.Col = 1
                WFecha2 = WVector1.Text
                WVector1.Col = 2
                WFactura2 = WVector1.Text
                WVector1.Col = 3
                WPrecio2 = WVector1.Text
                WVector1.Col = 4
                WCantidad2 = WVector1.Text
                    Else
                WFecha2 = ""
                WFactura2 = ""
                WPrecio2 = ""
                WCantidad2 = ""
            End If
        
            WVector1.Row = 3
            WVector1.Col = 4
            Auxi = Val(WVector1.Text)
    
            If Auxi <> 0 Then
                WVector1.Col = 1
                WFecha3 = WVector1.Text
                WVector1.Col = 2
                WFactura3 = WVector1.Text
                WVector1.Col = 3
                WPrecio3 = WVector1.Text
                WVector1.Col = 4
                WCantidad3 = WVector1.Text
                    Else
                WFecha3 = ""
                WFactura3 = ""
                WPrecio3 = ""
                WCantidad3 = ""
            End If
        
            WVector1.Row = 4
            WVector1.Col = 4
            Auxi = Val(WVector1.Text)
    
            If Auxi <> 0 Then
                WVector1.Col = 1
                WFecha4 = WVector1.Text
                WVector1.Col = 2
                WFactura4 = WVector1.Text
                WVector1.Col = 3
                WPrecio4 = WVector1.Text
                WVector1.Col = 4
                WCantidad4 = WVector1.Text
                    Else
                WFecha4 = ""
                WFactura4 = ""
                WPrecio4 = ""
                WCantidad4 = ""
            End If
        
            WVector1.Row = 5
            WVector1.Col = 4
            Auxi = Val(WVector1.Text)
    
            If Auxi <> 0 Then
                WVector1.Col = 1
                WFecha5 = WVector1.Text
                WVector1.Col = 2
                WFactura5 = WVector1.Text
                WVector1.Col = 3
                WPrecio5 = WVector1.Text
                WVector1.Col = 4
                WCantidad5 = WVector1.Text
                    Else
                WFecha5 = ""
                WFactura5 = ""
                WPrecio5 = ""
                WCantidad5 = ""
            End If
            Fecha.Caption = Date$
        
            If WPasa = "N" Then
                
                XParam = "'" + WClave + "','" + Cliente.Text + "','" + Terminado.Text + "','" + Precio.Text + "','" _
                         + Descripcion.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha.Caption + "','" + Pago.Text + "'"
                Set rstPrecios = db.OpenRecordset("AltaPrecios1 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WClave + "','" + Cliente.Text + "','" + Terminado.Text + "','" + Precio.Text + "','" _
                         + Descripcion.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha.Caption + "','" + Pago.Text + "'"
                Set rstPrecios = db.OpenRecordset("ModificaPrecios2 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            ZSql = ""
            ZSql = ZSql & "UPDATE Precios SET "
            ZSql = ZSql & "Estado = " + "'" + Str$(Estado.ListIndex) + "'"
            ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
            spPrecios = ZSql
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            
            If Val(Wempresa) = 1 Then
                XCodigo = Val(Mid$(Terminado.Text, 4, 5))
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    
                    ZDescripcion = Trim(DesTerminado.Caption) + " - " + Trim(Descripcion.Text)
                    ZDescripcionFarma = Descripcion.Text
                    
                    ZSql = ""
                    ZSql = ZSql & "UPDATE Precios SET "
                    ZSql = ZSql & "Descripcion = " + "'" + ZDescripcion + "',"
                    ZSql = ZSql & "DescripcionFarma = " + "'" + ZDescripcionFarma + "'"
                    ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
                    spPrecios = ZSql
                    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
                    
                End If
            End If
    
            Call CmdLimpiar_Click
            Call BotonPt_Click
            Cliente.SetFocus
            
        End If
        
    End If
    
End Sub

Private Sub cmdDelete_Click()
    If Cliente.Text <> "" And Terminado.Text <> "" Then
    
        Cliente.Text = UCase(Cliente.Text)
        Terminado.Text = UCase(Terminado.Text)
    
        WCliente = Cliente.Text
        WTerminado = Terminado.Text
        WClave = Cliente.Text + Terminado.Text
        
        spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
        Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
        If rstPrecios.RecordCount > 0 Then
            T$ = "Precios de Producto Terminado por Cliente"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spPrecios = "BorrarPrecios " + "'" + WClave + "'"
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
    End If
    Cliente.SetFocus
End Sub

Private Sub CmdLimpiar_Click()

    Cliente.Text = ""
    DesCliente.Caption = ""
    Terminado.Text = "  -     -   "
    Precio.Text = ""
    Descripcion.Text = ""
    DesTerminado.Caption = ""
    Fecha.Caption = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Estado.ListIndex = 0
    
    Cliente1.Text = ""
    DesCliente.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Precio1.Text = ""
    Fecha1.Caption = ""
    Pago1.Text = ""
    Despago1.Caption = ""
    Estado1.ListIndex = 0
    
    Call BotonPt_Click
    Cliente.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    Call CmdLimpiar_Click
    Cliente.SetFocus
    PrgPrecio.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub CmdLimpiar_gotFocus()
    Call CmdLimpiar_Click
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        WCliente = Cliente.Text
        spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!Razon
            rstCliente.Close
            Call Imprime_Datos
            Terminado.SetFocus
                Else
            Cliente.SetFocus
        End If
    End If
End Sub

Private Sub Command333_Click()

End Sub


Private Sub Mantenimiento_Click()

    If WGrabaNombre <> "S" Then
    
        Call Ingresa_claveII
        
            Else
            
        Rem WGrabaNombre = ""

        MenuNOmbre.Visible = True
        
    End If


End Sub

Private Sub Muestra4_Click()
    ZZDescripcion = Muestra4.TextMatrix(Muestra4.Row, 1)
    If Trim(ZZDescripcion) <> "" Then
        Descripcion.Text = ZZDescripcion
    End If
    Muestra4.Visible = False
End Sub


Private Sub Reproceso_Click()

    ZSql = "DELETE PreciosII"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)

    Erase ZZVector
    ZZLugar = 0
    ZZPasa = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Order by Terminado, Descripcion"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Left$(UCase(rstPrecios!Terminado), 2) = "PT" Or Left$(UCase(rstPrecios!Terminado), 2) = "PE" Then
                
                        If ZZPasa = 0 Then
                            ZZCorte = UCase(rstPrecios!Terminado)
                            ZZCorteII = Trim(rstPrecios!Descripcion)
                            ZZRenglon = 0
                            ZZPasa = 1
                        End If
                    
                        If ZZCorte <> UCase(rstPrecios!Terminado) Or ZZCorteII <> Trim(rstPrecios!Descripcion) Then
                        
                            If ZZCorte <> UCase(rstPrecios!Terminado) Then
                                ZZRenglon = 0
                            End If
                            
                            ZZLugar = ZZLugar + 1
                            ZZRenglon = ZZRenglon + 1
                            ZZVector(ZZLugar, 1) = ZZCorte
                            ZZVector(ZZLugar, 2) = ZZCorteII
                            ZZVector(ZZLugar, 3) = ZZRenglon
                    
                            ZZCorte = UCase(rstPrecios!Terminado)
                            ZZCorteII = Trim(rstPrecios!Descripcion)
                            
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
        
        ZZLugar = ZZLugar + 1
        ZZRenglon = ZZRenglon + 1
        ZZVector(ZZLugar, 1) = ZZCorte
        ZZVector(ZZLugar, 2) = ZZCorteII
        ZZVector(ZZLugar, 3) = ZZRenglon
        
    End If


    For ZZCiclo = 1 To ZZLugar
    
        ZZTerminado = ZZVector(ZZCiclo, 1)
        ZZDescripcion = ZZVector(ZZCiclo, 2)
        ZZRenglon = ZZVector(ZZCiclo, 3)
        
        Auxi = ZZRenglon
        Call Ceros(Auxi, 2)
        ZZClave = ZZTerminado + Auxi
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO PreciosII ("
        ZSql = ZSql & "Clave ,"
        ZSql = ZSql & "Terminado ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Nombre )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + ZZClave + "',"
        ZSql = ZSql & "'" + ZZTerminado + "',"
        ZSql = ZSql & "'" + ZZRenglon + "',"
        ZSql = ZSql & "'" + ZZDescripcion + "')"
    
        spPreciosII = ZSql
        Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
        
    Next ZZCiclo

    m$ = "Proceso Finalizado Correctamente"
    a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
    
    MenuNOmbre.Visible = False

End Sub

Private Sub Terminado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Terminado.Text = UCase(Terminado.Text)
        WTerminado = Terminado.Text
        spTerminado = "ConsultaTerminado " + "'" + Terminado.Text + "'"
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        If rstTerminado.RecordCount > 0 Then
            DesTerminado.Caption = rstTerminado!Descripcion
            Descripcion.Text = rstTerminado!Descripcion
            rstTerminado.Close
            If Left$(Terminado.Text, 2) <> "PT" And Left$(Terminado.Text, 2) <> "PE" And Left$(Terminado.Text, 2) <> "YQ" And Left$(Terminado.Text, 2) <> "YF" And Left$(Terminado.Text, 2) <> "YP" And Left$(Terminado.Text, 2) <> "YH" Then
                m$ = "Producto Terminado no esta autorizado para la venta"
                a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
                Exit Sub
            End If
            Label2.Caption = "Descripcion"
            Label14.Caption = ""
            If Val(Wempresa) = 1 Then
                XCodigo = Val(Mid$(Terminado.Text, 4, 5))
                If XCodigo >= 25000 And XCodigo <= 25999 Then
                    Label2.Caption = "Descripcion Adic."
                    Label14.Caption = "2do Renglon Etiquetas"
                End If
            End If
            
            Call Imprime_Datos
            Precio.SetFocus
                Else
            Terminado.SetFocus
        End If
    End If
End Sub

Private Sub Precio_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Descripcion.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Pago.SetFocus
    End If
    ''Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Descripcion_dblclick()

    Muestra4.Height = 4575
    Muestra4.Left = 200
    Muestra4.Top = 120
    Muestra4.Width = 6700
    
    Muestra4.Clear
    Muestra4.Row = 0
    
    Muestra4.Col = 1
    Muestra4.Text = "Nombre"
    
    
    Muestra4.ColWidth(0) = 100
    Muestra4.ColWidth(1) = 6000
    
    Muestra4.Visible = True
    
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM PreciosII"
    ZSql = ZSql + " Where PreciosII.Terminado = " + "'" + Terminado.Text + "'"
    ZSql = ZSql + " Order by Clave"
    spPreciosII = ZSql
    Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosII.RecordCount > 0 Then
    
        With rstPreciosII
            .MoveFirst
            Do
                If .EOF = False Then
    
                    Renglon = Renglon + 1
                    
                    Muestra4.TextMatrix(Renglon, 1) = rstPreciosII!Nombre
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPreciosII.Close
    End If
    
    Muestra4.Row = 1
    Muestra4.Col = 1
    Muestra4.TopRow = 1

End Sub

Private Sub Pago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Pago.Text) <> 0 Then
            spPago = "ConsultaPago " + "'" + Pago.Text + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                DesPago.Caption = rstPago!Nombre
                rstPago.Close
                Precio.SetFocus
            End If
                Else
            DesPago.Caption = ""
            Precio.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.Text = UCase(DesdeCliente.Text)
        Hastacliente.Text = DesdeCliente.Text
        Hastacliente.SetFocus
    End If
End Sub

Private Sub HastaCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastacliente.Text = UCase(Hastacliente.Text)
        DesdeTerminado.SetFocus
    End If
End Sub

Private Sub DesdeTerminado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeTerminado.Text = UCase(DesdeTerminado.Text)
        HastaTerminado.SetFocus
    End If
End Sub

Private Sub HastaTerminado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaTerminado.Text = UCase(HastaTerminado.Text)
        ListaVendedor.SetFocus
    End If
End Sub

Private Sub ListaVendedor_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.SetFocus
    End If
End Sub

Private Sub Consulta_Click()

     Opcion.Clear
     
     Opcion.AddItem "Precios de Producto Terminado por Cliente"
     Opcion.AddItem "Clientes"
     Opcion.AddItem "Productos Terminados"
     Opcion.AddItem "Condiciones de Pago"

     Opcion.Visible = True
     
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    'XIndice = 0
    
    Select Case XIndice
        Case 0
            spPrecios = "ListaPrecios"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
            
                With rstPrecios
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstPrecios!Cliente = Cliente.Text Then
                                IngresaItem = rstPrecios!Cliente + " " + rstPrecios!Terminado
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstPrecios!Clave
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPrecios.Close
            End If
            
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
            
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
        
        Case 2
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If Left$(rstTerminado!Codigo, 2) = "PT" Or Left$(rstTerminado!Codigo, 2) = "PE" Then
                                IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                Pantalla.AddItem IngresaItem
                                IngresaItem = rstTerminado!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
            
        Case 3
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
            
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstPago!Pago) + " " + rstPago!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstPago!Pago
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPago.Close
            End If
            Ayuda.Text = ""
            Ayuda.Visible = True
            Ayuda.SetFocus
        
        Case Else
    End Select
            
    Pantalla.Visible = True

End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            DA = Len(rstCliente!Razon) - WEspacios
                
                            For aa = 1 To DA
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                                    Auxi = rstCliente!Cliente
                                    IngresaItem = Auxi + "    " + rstCliente!Razon
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCliente!Cliente
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 2
            spTerminado = "ListaTerminado"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                With rstTerminado
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            If Left$(rstTerminado!Codigo, 2) = "PT" Or Left$(rstTerminado!Codigo, 2) = "PE" Then
                                DA = Len(rstTerminado!Descripcion) - WEspacios
                
                                For aa = 1 To DA
                                    If Left$(Ayuda.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                        Auxi = rstTerminado!Codigo
                                        IngresaItem = Auxi + "    " + rstTerminado!Descripcion
                                        Pantalla.AddItem IngresaItem
                                        IngresaItem = rstTerminado!Codigo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                            End If
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstTerminado.Close
            End If
            
        Case 3
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            DA = Len(rstPago!Nombre) - WEspacios
                
                            For aa = 1 To DA
                                If Left$(Ayuda.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                    Auxi = Str$(rstPago!Pago)
                                    IngresaItem = Auxi + "    " + rstPago!Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstPago!Pago
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstPago.Close
            End If
            
        Case Else
        
    End Select
            
    End If

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            WClave = WIndice.List(Indice)
            spPrecios = "ConsultaPrecios " + "'" + WClave + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                Cliente.Text = rstPrecios!Cliente
                Terminado.Text = rstPrecios!Terminado
                rstPrecios.Close
                Call Imprime_Datos
                    Else
                CmdLimpiar_Click
            End If
            Precio.SetFocus
        
        Case 1
            Indice = Pantalla.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente.Text = rstCliente!Cliente
                rstCliente.Close
                Call Imprime_Datos
                    Else
                Cliente.Text = WCliente
            End If
            Cliente.SetFocus
            
        Case 2
            Indice = Pantalla.ListIndex
            WTerminado = WIndice.List(Indice)
            spTerminado = "ConsultaTerminado " + "'" + WTerminado + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                Terminado.Text = rstTerminado!Codigo
                rstTerminado.Close
                Call Imprime_Datos
                    Else
                Terminado.Text = WTerminado
            End If
            Terminado.SetFocus
            
        Case 3
            Indice = Pantalla.ListIndex
            WPago = WIndice.List(Indice)
            spPago = "ConsultaPago " + "'" + WPago + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                Pago.Text = rstPago!Pago
                rstPago.Close
                Call Imprime_Descripcion
                    Else
                Pago.Text = WPago
            End If
            Pago.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Lista_Click()
    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeTerminado.Text = "  -     -   "
    HastaTerminado.Text = "  -     -   "
    ListaVendedor.Text = ""
    Panta.Value = False
    Impresora.Value = True
    Frame2.Visible = True
    DesdeCliente.SetFocus
End Sub

Private Sub Primer_Click()

    On Error GoTo WError
    
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Cliente.Text = rstPrecios!Cliente
            Terminado.Text = rstPrecios!Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Ultimo_Click()

   On Error GoTo Error_ultimo
    
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveLast
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Precios de Prodcuto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
 End Sub

Private Sub Anterior_Click()

    On Error GoTo WError
    
    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "AnteriorPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveLast
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Precio.SetFocus
    
End Sub

Private Sub Siguiente_Click()

    On Error GoTo WError
    
    Cliente.Text = UCase(Cliente.Text)
    Terminado.Text = UCase(Terminado.Text)
    
    WCliente = Cliente.Text
    WTerminado = Terminado.Text
    WClave = Cliente.Text + Terminado.Text
    
    spPrecios = "PosteriorPrecios " + "'" + WClave + "'"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
        With rstPrecios
            .MoveFirst
            Cliente.Text = !Cliente
            Terminado.Text = !Terminado
            rstPrecios.Close
            Call Imprime_Datos
        End With
    End If
    
    Cliente.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Producto Terminado", "No existe registro en el archivo")
     Call CmdLimpiar_Click
     Cliente.SetFocus
    
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear
    WVector1.Font.Bold = True
    
    WVector1.FixedCols = 1
    WVector1.Cols = 5
    WVector1.FixedRows = 1
    WVector1.Rows = 101
    
    WVector1.ColWidth(0) = 200
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Fecha"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector1.Text = "Factura"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector1.Text = "Precio Unitario"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector1.Text = "Cantidad"
                WVector1.ColWidth(Ciclo) = 1500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        WTitulo(Ciclo).Text = WVector1.Text
        WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        WTitulo(Ciclo).Width = WVector1.CellWidth
        WTitulo(Ciclo).Height = WVector1.CellHeight
        WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
    OPEN_FILE_Empresa
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(Wempresa)
        If .NoMatch = False Then
            PrgPrecio.Caption = "Ingreso de Precios por Cliente :  " + !Nombre
        End If
    End With
End Sub

Private Sub Acepta1_Click()

    DesdeArticulo.Text = UCase(DesdeArticulo.Text)
    HastaArticulo.Text = UCase(HastaArticulo.Text)
    DesdeCliente1.Text = UCase(DesdeCliente1.Text)
    HastaCliente1.Text = UCase(HastaCliente1.Text)
    
    Listado.WindowTitle = "Listado de Precios de Maetrias Primas por Cliente"
    Listado.WindowTop = 0
    Listado.WindowLeft = 0
    Listado.WindowWidth = Screen.Width
    Listado.WindowHeight = Screen.Height

    If Impresora1.Value = True Then
        Listado.Destination = 1
            Else
        Listado.Destination = 0
    End If
    
    If Val(ListaVendedor1.Text) <> 0 Then
        WDesdeVendedor = ListaVendedor1.Text
        WHastaVendedor = ListaVendedor1.Text
            Else
        WDesdeVendedor = "0"
        WHastaVendedor = "9999"
    End If
    
    Uno = "{PreciosMp.Articulo} in " + Chr$(34) + DesdeArticulo.Text + Chr$(34) + " to " + Chr$(34) + HastaArticulo.Text + Chr$(34)
    Dos = " and " + "{PreciosMp.Cliente} in " + Chr$(34) + DesdeCliente1.Text + Chr$(34) + " to " + Chr$(34) + HastaCliente1.Text + Chr$(34)
    Tres = " and " + "{Cliente.Vendedor} in " + WDesdeVendedor + " to " + WHastaVendedor
    
    Listado.GroupSelectionFormula = Uno + Dos + Tres
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Listado.SQLQuery = "SELECT PreciosMp.Cliente, PreciosMp.Articulo, PreciosMp.Precio, " _
                    + "Cliente.Razon, Cliente.Vendedor, " _
                    + "Articulo.Descripcion " _
                    + "From " _
                    + DSQ + ".dbo.PreciosMp PreciosMp, " _
                    + DSQ + ".dbo.Cliente Cliente, " _
                    + DSQ + ".dbo.Articulo Articulo " _
                    + "Where " _
                    + "PreciosMp.Cliente = Cliente.Cliente AND " _
                    + "PreciosMp.Articulo = Articulo.Codigo AND " _
                    + "PreciosMp.Cliente >= '" + DesdeCliente1.Text + "' AND " _
                    + "PreciosMp.Cliente <= '" + HastaCliente1.Text + "' AND " _
                    + "PreciosMp.Articulo >= '" + DesdeArticulo.Text + "' AND " _
                    + "PreciosMp.Articulo <= '" + HastaArticulo.Text + "'"
    
    Listado.DataFiles(1) = Wempresa + "auxi.mdb"
    Listado.Connect = Connect()
    Listado.ReportFileName = "WPreciosMp.rpt"
    
    Cliente1.SetFocus
    Listado.Action = 1
    Frame5.Visible = False
    
End Sub

Private Sub Cancela1_click()
    Frame5.Visible = False
End Sub

Sub Imprime_Descripcion1()

    Rem lee Cliente

    spCliente = "ConsultaCliente " + "'" + Cliente1.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesCliente1.Caption = rstCliente!Razon
        rstCliente.Close
            Else
        DesCliente1.Caption = ""
    End If
    
    Rem lee articulo
    
    spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        DesArticulo.Caption = rstArticulo!Descripcion
        rstArticulo.Close
            Else
        DesArticulo.Caption = ""
    End If
    
    spPago = "ConsultaPago " + "'" + Pago1.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        Despago1.Caption = rstPago!Nombre
        rstPago.Close
            Else
        Despago1.Caption = ""
    End If
    
End Sub

Sub Verifica_datos1()
    If Val(Precio1.Text) = 0 Then
        Precio1.Text = "0"
    End If
    If Val(Pago1.Text) = 0 Then
        Pago1.Text = "0"
    End If
End Sub

Sub Format_datos1()
    Precio1.Text = Pusing("###,###.##", Precio1.Text)
End Sub

Sub Imprime_Datos1()

    Cliente1.Text = UCase(Cliente1.Text)
    Articulo.Text = UCase(Articulo.Text)
    
    WCliente = Cliente1.Text
    WArticulo = Articulo.Text
    WClave = Cliente1.Text + Articulo.Text
    
    spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        Cliente1.Text = rstPreciosMp!Cliente
        Articulo.Text = rstPreciosMp!Articulo
        Precio1.Text = rstPreciosMp!Precio
        Fecha1.Caption = IIf(IsNull(rstPreciosMp!Fecha), "", rstPreciosMp!Fecha)
        Pago1.Text = IIf(IsNull(rstPreciosMp!Pago), "0", rstPreciosMp!Pago)
        ZEstado = IIf(IsNull(rstPreciosMp!Estado), "0", rstPreciosMp!Estado)
        Estado1.ListIndex = ZEstado
        Call Format_datos1
            
        'columna 1
        
        Call Limpia_Vector1
                    
        WVector2.Row = 1
    
        If rstPreciosMp!Cantidad1 <> 0 Then
            WVector2.Col = 1
            WVector2.Text = rstPreciosMp!Fecha1
            WVector2.Col = 2
            WVector2.Text = rstPreciosMp!Factura1
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Precio1)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Cantidad1)
                Else
            WVector2.Col = 1
            WVector2.Text = ""
            WVector2.Col = 2
            WVector2.Text = ""
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", Dada)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 2
    
        WVector2.Row = 2
            
        If rstPreciosMp!Cantidad2 <> 0 Then
            WVector2.Col = 1
            WVector2.Text = rstPreciosMp!Fecha2
            WVector2.Col = 2
            WVector2.Text = rstPreciosMp!Factura2
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Precio2)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Cantidad2)
                Else
            WVector2.Col = 1
            WVector2.Text = ""
            WVector2.Col = 2
            WVector2.Text = ""
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", Dada)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 3
    
        WVector2.Row = 3
            
        If rstPreciosMp!Cantidad3 <> 0 Then
            WVector2.Col = 1
            WVector2.Text = rstPreciosMp!Fecha3
            WVector2.Col = 2
            WVector2.Text = rstPreciosMp!Factura3
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Precio3)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Cantidad3)
                Else
            WVector2.Col = 1
            WVector2.Text = ""
            WVector2.Col = 2
            WVector2.Text = ""
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", Dada)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 4
    
        WVector2.Row = 4
            
        If rstPreciosMp!Cantidad4 <> 0 Then
            WVector2.Col = 1
            WVector2.Text = rstPreciosMp!Fecha4
            WVector2.Col = 2
            WVector2.Text = rstPreciosMp!Factura4
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Precio4)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Cantidad4)
                Else
            WVector2.Col = 1
            WVector2.Text = ""
            WVector2.Col = 2
            WVector2.Text = ""
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", Dada)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", Dada)
        End If
    
        'columna 5
    
        WVector2.Row = 5
        
        If rstPreciosMp!Cantidad5 <> 0 Then
            WVector2.Col = 1
            WVector2.Text = rstPreciosMp!Fecha5
            WVector2.Col = 2
            WVector2.Text = rstPreciosMp!Factura5
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Precio5)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", rstPreciosMp!Cantidad5)
                Else
            WVector2.Col = 1
            WVector2.Text = ""
            WVector2.Col = 2
            WVector2.Text = ""
            WVector2.Col = 3
            WVector2.Text = Pusing("###,###.##", Dada)
            WVector2.Col = 4
            WVector2.Text = Pusing("###,###.##", Dada)
        End If
        rstPreciosMp.Close

    End If
    
    Call Imprime_Descripcion1
    
End Sub

Private Sub cmdAdd1_Click()

    If WGraba <> "S" Then
    
        WGrabaProceso = 1
        Call Ingresa_clave
        
            Else
            
        WGraba = ""
        
        WReventa = 0
        Articulo.Text = UCase(Articulo.Text)
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
            rstArticulo.Close
        End If
        If WReventa = 0 Then
            m$ = "Articulo no esta autorizado para la reventa"
            a% = MsgBox(m$, 0, "Ingreso de Precios de Materias Primas")
            Exit Sub
        End If

        spArticulo = "ConsultaArticulo " + "'" + UCase(Articulo.Text) + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            rstArticulo.Close
                Else
            m$ = "Articulo no esta ingresado"
            a% = MsgBox(m$, 0, "Ingreso de Precios de Materias Primas")
            Exit Sub
        End If

        If Cliente1.Text <> "" And Articulo.Text <> "" Then
    
            Cliente1.Text = UCase(Cliente1.Text)
            Articulo.Text = UCase(Articulo.Text)
    
            WCliente1 = Cliente1.Text
            WArticulo = Articulo.Text
            WClave = Cliente1.Text + Articulo.Text
            
            ZImpreVto = 0
            spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                ZImpreVto = IIf(IsNull(rstCliente!ImpreVto), "0", rstCliente!ImpreVto)
                rstCliente.Close
            End If
            
            If ZImpreVto = 1 Then
            
                ZVida = 0
                spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    ZVida = IIf(IsNull(rstArticulo!Meses), "0", rstArticulo!Meses)
                    rstArticulo.Close
                End If
                
                If ZVida = 0 Then
                    m$ = "Atencion: El producto terminado no posee vida util y el cliente lo exige"
                    a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
                End If
            
            End If
            
            spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                WPasa = "S"
                    Else
                WPasa = "N"
            End If
        
            Call Verifica_datos1
        
            WVector2.Row = 1
            WVector2.Col = 4
            Auxi = Val(WVector2.Text)
    
            If Auxi <> 0 Then
                WVector2.Col = 1
                WFecha1 = WVector2.Text
                WVector2.Col = 2
                WFactura1 = WVector2.Text
                WVector2.Col = 3
                WPrecio2 = WVector2.Text
                WVector2.Col = 4
                WCantidad1 = WVector2.Text
                    Else
                WFecha1 = ""
                WFactura1 = ""
                WPrecio1 = ""
                WCantidad1 = ""
            End If
        
            WVector2.Row = 2
            WVector2.Col = 4
            Auxi = Val(WVector2.Text)
    
            If Auxi <> 0 Then
                WVector2.Col = 1
                WFecha2 = WVector2.Text
                WVector2.Col = 2
                WFactura2 = WVector2.Text
                WVector2.Col = 3
                WPrecio2 = WVector2.Text
                WVector2.Col = 4
                WCantidad2 = WVector2.Text
                    Else
                WFecha2 = ""
                WFactura2 = ""
                WPrecio2 = ""
                WCantidad2 = ""
            End If
        
            WVector2.Row = 3
            WVector2.Col = 4
            Auxi = Val(WVector2.Text)
    
            If Auxi <> 0 Then
                WVector2.Col = 1
                WFecha3 = WVector2.Text
                WVector2.Col = 2
                WFactura3 = WVector2.Text
                WVector2.Col = 3
                WPrecio3 = WVector2.Text
                WVector2.Col = 4
                WCantidad3 = WVector2.Text
                    Else
                WFecha3 = ""
                WFactura3 = ""
                WPrecio3 = ""
                WCantidad3 = ""
            End If
        
            WVector2.Row = 4
            WVector2.Col = 4
            Auxi = Val(WVector2.Text)
    
            If Auxi <> 0 Then
                WVector2.Col = 1
                WFecha4 = WVector2.Text
                WVector2.Col = 2
                WFactura4 = WVector2.Text
                WVector2.Col = 3
                WPrecio4 = WVector2.Text
                WVector2.Col = 4
                WCantidad4 = WVector2.Text
                    Else
                WFecha4 = ""
                WFactura4 = ""
                WPrecio4 = ""
                WCantidad4 = ""
            End If
        
            WVector2.Row = 5
            WVector2.Col = 4
            Auxi = Val(WVector2.Text)
    
            If Auxi <> 0 Then
                WVector2.Col = 1
                WFecha5 = WVector2.Text
                WVector2.Col = 2
                WFactura5 = WVector2.Text
                WVector2.Col = 3
                WPrecio5 = WVector2.Text
                WVector2.Col = 4
                WCantidad5 = WVector2.Text
                    Else
                WFecha5 = ""
                WFactura5 = ""
                WPrecio5 = ""
                WCantidad5 = ""
            End If
            Fecha1.Caption = Date$
        
            If WPasa = "N" Then
                XParam = "'" + WClave + "','" + Cliente1.Text + "','" + Articulo.Text + "','" + Precio1.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha1.Caption + "','" + Pago1.Text + "'"
                Set rstPreciosMp = db.OpenRecordset("AltaPreciosMp " + XParam, dbOpenSnapshot, dbSQLPassThrough)
                    Else
                XParam = "'" + WClave + "','" + Cliente1.Text + "','" + Articulo.Text + "','" + Precio1.Text + "','" _
                         + WFecha1 + "','" + WFactura1 + "','" + WPrecio1 + "','" + WCantidad1 + "','" _
                         + WFecha2 + "','" + WFactura2 + "','" + WPrecio2 + "','" + WCantidad2 + "','" _
                         + WFecha3 + "','" + WFactura3 + "','" + WPrecio3 + "','" + WCantidad3 + "','" _
                         + WFecha4 + "','" + WFactura4 + "','" + WPrecio4 + "','" + WCantidad4 + "','" _
                         + WFecha5 + "','" + WFactura5 + "','" + WPrecio5 + "','" + WCantidad5 + "','" _
                         + Date$ + "','" + Fecha1.Caption + "','" + Pago1.Text + "'"
                Set rstPreciosMp = db.OpenRecordset("ModificaPreciosMp " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
    
            ZSql = ""
            ZSql = ZSql & "UPDATE PreciosMp SET "
            ZSql = ZSql & "Estado = " + "'" + Str$(Estado1.ListIndex) + "'"
            ZSql = ZSql & " Where Clave = " + "'" + WClave + "'"
            spPreciosMp = ZSql
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    
            Call CmdLimpiar1_Click
            Cliente1.SetFocus
            
        End If
        
    End If
    
End Sub

Private Sub cmdDelete1_Click()
    If Cliente1.Text <> "" And Articulo.Text <> "" Then
    
        Cliente1.Text = UCase(Cliente1.Text)
        Articulo.Text = UCase(Articulo.Text)
    
        WCliente = Cliente1.Text
        WArticulo = Articulo.Text
        WClave = Cliente1.Text + Articulo.Text
        
        spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
        Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
        If rstPreciosMp.RecordCount > 0 Then
            T$ = "Precios de Materias Primas por Cliente"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                spPreciosMp = "BorrarPreciosMp " + "'" + WClave + "'"
                Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenDynaset, dbSQLPassThrough)
                Call CmdLimpiar1_Click
            End If
        End If
        
    End If
    Cliente1.SetFocus
End Sub

Private Sub Cliente1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente1.Text = UCase(Cliente1.Text)
        WCliente = Cliente1.Text
        spCliente = "ConsultaCliente " + "'" + Cliente1.Text + "'"
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente1.Caption = rstCliente!Razon
            rstCliente.Close
            Call Imprime_Datos1
            Articulo.SetFocus
                Else
            Cliente1.SetFocus
        End If
    End If
End Sub

Private Sub Articulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Articulo.Text = UCase(Articulo.Text)
        WArticulo = Articulo.Text
        spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
            rstArticulo.Close
            If WReventa = 0 Then
                m$ = "Articulo no esta autorizado para la reventa"
                a% = MsgBox(m$, 0, "Ingreso de Precios de Materias Primas")
                Exit Sub
            End If
            Call Imprime_Datos1
            Precio1.SetFocus
                Else
            Articulo.SetFocus
        End If
    End If
End Sub

Private Sub Precio1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
          Pago1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Pago1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Pago1.Text) <> 0 Then
            spPago = "ConsultaPago " + "'" + Pago1.Text + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                Despago1.Caption = rstPago!Nombre
                rstPago.Close
                Precio1.SetFocus
            End If
                Else
            Despago1.Caption = ""
            Precio1.SetFocus
        End If
    End If
End Sub

Private Sub DesdeCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente1.Text = UCase(DesdeCliente1.Text)
        HastaCliente1.Text = DesdeCliente1.Text
        HastaCliente1.SetFocus
    End If
End Sub

Private Sub HastaCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCliente1.Text = UCase(HastaCliente1.Text)
        DesdeArticulo.SetFocus
    End If
End Sub

Private Sub DesdeArticulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArticulo.Text = UCase(DesdeArticulo.Text)
        HastaArticulo.SetFocus
    End If
End Sub

Private Sub HastaArticulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArticulo.Text = UCase(HastaArticulo.Text)
        ListaVendedor1.SetFocus
    End If
End Sub

Private Sub ListaVendedor1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente1.SetFocus
    End If
End Sub

Private Sub Consulta1_Click()

     Opcion1.Clear
     
     Opcion1.AddItem "Precios de Materias Primas por Cliente"
     Opcion1.AddItem "Clientes"
     Opcion1.AddItem "Materias Primas"
     Opcion1.AddItem "Condiciones de Pago"

     Opcion1.Visible = True
     
End Sub

Private Sub Opcion1_Click()

    Opcion1.Visible = False
     
    Dim IngresaItem As String

    Pantalla1.Clear
    WIndice.Clear

    XIndice = Opcion1.ListIndex
    'XIndice = 0
    
    Select Case XIndice
        Case 0
            spPreciosMp = "ListaPreciosMp"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
            
                With rstPreciosMp
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If rstPreciosMp!Cliente = Cliente1.Text Then
                                IngresaItem = rstPreciosMp!Cliente + " " + rstPreciosMp!Articulo
                                Pantalla1.AddItem IngresaItem
                                IngresaItem = rstPreciosMp!Clave
                                WIndice.AddItem IngresaItem
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPreciosMp.Close
            End If
            
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
            
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstCliente!Cliente + " " + rstCliente!Razon
                            Pantalla1.AddItem IngresaItem
                            IngresaItem = rstCliente!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            Ayuda1.Text = ""
            Ayuda1.Visible = True
            Ayuda1.SetFocus
        
        Case 2
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                            If WReventa = 1 Then
                                IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                Pantalla1.AddItem IngresaItem
                                IngresaItem = rstArticulo!Codigo
                                WIndice.AddItem IngresaItem
                            End If
                                .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            Ayuda1.Text = ""
            Ayuda1.Visible = True
            Ayuda1.SetFocus
            
        Case 3
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
            
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstPago!Pago) + " " + rstPago!Nombre
                            Pantalla1.AddItem IngresaItem
                            IngresaItem = rstPago!Pago
                            WIndice.AddItem IngresaItem
                                .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstPago.Close
            End If
            Ayuda1.Text = ""
            Ayuda1.Visible = True
            Ayuda1.SetFocus
        
        Case Else
    End Select
            
    Pantalla1.Visible = True

End Sub

Private Sub Ayuda1_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla1.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda1.Text)
    
    Select Case XIndice
        Case 1
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                With rstCliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            DA = Len(rstCliente!Razon) - WEspacios
                
                            For aa = 1 To DA
                                If Left$(Ayuda1.Text, WEspacios) = Mid$(!Razon, aa, WEspacios) Then
                                    Auxi = rstCliente!Cliente
                                    IngresaItem = Auxi + "    " + rstCliente!Razon
                                    Pantalla1.AddItem IngresaItem
                                    IngresaItem = rstCliente!Cliente
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstCliente.Close
            End If
            
        Case 2
            spArticulo = "ListaArticulo"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            WReventa = IIf(IsNull(rstArticulo!Reventa), "0", rstArticulo!Reventa)
                            If WReventa = 1 Then
                                DA = Len(rstArticulo!Descripcion) - WEspacios
                
                                For aa = 1 To DA
                                    If Left$(Ayuda1.Text, WEspacios) = Mid$(!Descripcion, aa, WEspacios) Then
                                        Auxi = rstArticulo!Codigo
                                        IngresaItem = Auxi + "    " + rstArticulo!Descripcion
                                        Pantalla1.AddItem IngresaItem
                                        IngresaItem = rstArticulo!Codigo
                                        WIndice.AddItem IngresaItem
                                        Exit For
                                    End If
                                Next aa
                            End If
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstArticulo.Close
            End If
            
        Case 3
            spPago = "ListaPago"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                With rstPago
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            DA = Len(rstPago!Nombre) - WEspacios
                
                            For aa = 1 To DA
                                If Left$(Ayuda1.Text, WEspacios) = Mid$(!Nombre, aa, WEspacios) Then
                                    Auxi = Str$(rstPago!Pago)
                                    IngresaItem = Auxi + "    " + rstPago!Nombre
                                    Pantalla1.AddItem IngresaItem
                                    IngresaItem = rstPago!Pago
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next aa
                            .MoveNext
                            
                                Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
                rstPago.Close
            End If
            
        Case Else
        
    End Select
            
    End If

End Sub


Private Sub Pantalla1_Click()
    Pantalla1.Visible = False
    Ayuda1.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla1.ListIndex
            WClave = WIndice.List(Indice)
            spPreciosMp = "ConsultaPreciosMp " + "'" + WClave + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                Cliente1.Text = rstPreciosMp!Cliente
                Articulo.Text = rstPreciosMp!Articulo
                rstPreciosMp.Close
                Call Imprime_Datos1
                    Else
                CmdLimpiar1_Click
            End If
            Precio1.SetFocus
        
        Case 1
            Indice = Pantalla1.ListIndex
            WCliente = WIndice.List(Indice)
            spCliente = "ConsultaCliente " + "'" + WCliente + "'"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                Cliente1.Text = rstCliente!Cliente
                rstCliente.Close
                Call Imprime_Datos1
                    Else
                Cliente1.Text = WCliente
            End If
            Cliente1.SetFocus
            
        Case 2
            Indice = Pantalla1.ListIndex
            WArticulo = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + WArticulo + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Articulo.Text = rstArticulo!Codigo
                rstArticulo.Close
                Call Imprime_Datos1
                    Else
                Articulo.Text = WArticulo
            End If
            Articulo.SetFocus
            
        Case 3
            Indice = Pantalla1.ListIndex
            WPago = WIndice.List(Indice)
            spPago = "ConsultaPago " + "'" + WPago + "'"
            Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
            If rstPago.RecordCount > 0 Then
                Pago1.Text = rstPago!Pago
                rstPago.Close
                Call Imprime_Descripcion1
                    Else
                Pago1.Text = WPago
            End If
            Pago1.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub Lista1_Click()
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArticulo.Text = "  -   -   "
    HastaArticulo.Text = "  -   -   "
    ListaVendedor1.Text = ""
    Panta1.Value = False
    Impresora1.Value = True
    Frame5.Visible = True
    DesdeCliente1.SetFocus
End Sub

Private Sub Primer1_Click()

    On Error GoTo WError
    
    spPreciosMp = "ListaPreciosMp"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        With rstPreciosMp
            .MoveFirst
            Cliente1.Text = rstPreciosMp!Cliente
            Articulo.Text = rstPreciosMp!Articulo
            rstPreciosMp.Close
            Call Imprime_Datos1
        End With
    End If
    
    Cliente1.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Materias Primas", "No existe registro en el archivo")
     Call CmdLimpiar1_Click
     Cliente1.SetFocus
 End Sub

Private Sub Ultimo1_Click()

   On Error GoTo Error_ultimo
    
    spPreciosMp = "ListaPreciosMp"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        With rstPreciosMp
            .MoveLast
            Cliente1.Text = !Cliente
            Articulo.Text = !Articulo
            rstPreciosMp.Close
            Call Imprime_Datos1
        End With
    End If
    
    Cliente1.SetFocus
    
    Exit Sub
    
Error_ultimo:
     coderr = Err
     Call Errores(coderr, "Precios de Materias Primas", "No existe registro en el archivo")
     Call CmdLimpiar1_Click
     Cliente1.SetFocus
 End Sub

Private Sub Anterior1_Click()

    On Error GoTo WError
    
    Cliente1.Text = UCase(Cliente1.Text)
    Articulo.Text = UCase(Articulo.Text)
    
    WCliente = Cliente1.Text
    WArticulo = Articulo.Text
    WClave = Cliente1.Text + Articulo.Text
    
    spPreciosMp = "AnteriorPreciosMp " + "'" + WClave + "'"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        With rstPreciosMp
            .MoveLast
            Cliente1.Text = !Cliente
            Articulo.Text = !Articulo
            rstPreciosMp.Close
            Call Imprime_Datos1
        End With
    End If
    
    Cliente1.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Materias Primas", "No existe registro en el archivo")
     Call CmdLimpiar1_Click
     Precio1.SetFocus
    
End Sub

Private Sub Siguiente1_Click()

    On Error GoTo WError
    
    Cliente1.Text = UCase(Cliente1.Text)
    Articulo.Text = UCase(Articulo.Text)
    
    WCliente = Cliente1.Text
    WArticulo = Articulo.Text
    WClave = Cliente1.Text + Articulo.Text
    
    spPreciosMp = "PosteriorPreciosMp " + "'" + WClave + "'"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    If rstPreciosMp.RecordCount > 0 Then
        With rstPreciosMp
            .MoveFirst
            Cliente1.Text = !Cliente
            Articulo.Text = !Articulo
            rstPreciosMp.Close
            Call Imprime_Datos1
        End With
    End If
    
    Cliente1.SetFocus
    
    Exit Sub

WError:
     coderr = Err
     Call Errores(coderr, "Precios de Materias Primas", "No existe registro en el archivo")
     Call CmdLimpiar1_Click
     Cliente1.SetFocus
    
End Sub

Private Sub Limpia_Vector1()

    WVector2.Clear
    WVector2.Font.Bold = True
    
    WVector2.FixedCols = 1
    WVector2.Cols = 5
    WVector2.FixedRows = 1
    WVector2.Rows = 101
    
    WVector2.ColWidth(0) = 200
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector2.Text = "Fecha"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 2
                WVector2.Text = "Factura"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 3
                WVector2.Text = "Precio Unitario"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
            Case 4
                WVector2.Text = "Cantidad"
                WVector2.ColWidth(Ciclo) = 1500
                WVector2.ColAlignment(Ciclo) = flexAlignRightCenter
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector2.Row = 0
    For Ciclo = 1 To WVector2.Cols - 1
        WVector2.Col = Ciclo
        WTitulo1(Ciclo).Text = WVector2.Text
        WTitulo1(Ciclo).Left = WVector2.CellLeft + WVector2.Left
        WTitulo1(Ciclo).Top = WVector2.CellTop + WVector2.Top
        WTitulo1(Ciclo).Width = WVector2.CellWidth
        WTitulo1(Ciclo).Height = WVector2.CellHeight
        WTitulo1(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA GRILLA
    
    WAncho = 400
    For Ciclo = 0 To WVector2.Cols - 1
        WAncho = WAncho + WVector2.ColWidth(Ciclo)
    Next Ciclo
    WVector2.Width = WAncho

    ' Size the columns.
    Font.Name = WVector2.Font.Name
    Font.Size = WVector2.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector2.AllowUserResizing = flexResizeBoth
    
    WVector2.Col = 1
    WVector2.Row = 1
    
End Sub

Private Sub CmdLimpiar1_Click()

    Cliente.Text = ""
    DesCliente.Caption = ""
    Terminado.Text = "  -     -   "
    Precio.Text = ""
    Descripcion.Text = ""
    DesTerminado.Caption = ""
    Fecha.Caption = ""
    Pago.Text = ""
    DesPago.Caption = ""
    Estado.ListIndex = 0
    
    Cliente1.Text = ""
    DesCliente1.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Precio1.Text = ""
    Fecha1.Caption = ""
    Pago1.Text = ""
    Despago1.Caption = ""
    Estado1.ListIndex = 0
    
    Call BotonDy_Click
    Cliente1.SetFocus
    
End Sub

Private Sub cmdClose1_Click()
    Call CmdLimpiar1_Click
    PrgPrecio.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Load()

    Estado.Clear
    
    Estado.AddItem "Activo"
    Estado.AddItem "Historico"
    Estado.AddItem "Cotizacion"
    
    Estado.ListIndex = 0
    
    Estado1.Clear
    
    Estado1.AddItem "Activo"
    Estado1.AddItem "Historico"
    Estado1.AddItem "Cotizacion"
    
    Estado1.ListIndex = 0

    PantaDy.Visible = False
    PantaPt.Height = 7575
    PantaPt.Left = 120
    PantaPt.Top = 360
    PantaPt.Width = 9855
    PantaPt.Visible = True
    
    Terminado.Text = "  -     -   "
    DesTerminado.Caption = ""
    Cliente.Text = ""
    DesCliente.Caption = ""
    Precio.Text = ""
    Descripcion.Text = ""
    Fecha.Caption = ""
    
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Cliente1.Text = ""
    DesCliente1.Caption = ""
    Precio1.Text = ""
    Fecha1.Caption = ""
    
    Call Limpia_Vector
    
End Sub

Sub Ingresa_clave()

    WClave.Text = ""
    Clave1.Visible = True
    WClave.SetFocus
    
End Sub

Private Sub CancelaGraba_Click()

    Clave1.Visible = False

End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGraba = "N"
        WClave.Text = UCase(WClave.Text)
        If WClave.Text = "CHANGE" Then
            WGraba = "S"
            Clave1.Visible = False
            If WGrabaProceso = 0 Then
                Call cmdAdd_Click
                    Else
                Call cmdAdd1_Click
            End If
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Ingreso de Precios de Productos Terminados")
            WClave.SetFocus
        End If
    End If

End Sub








Sub Ingresa_claveII()

    WClaveNombre.Text = ""
    PantaClaveNombre.Visible = True
    WClaveNombre.SetFocus
    
End Sub

Private Sub CancelaGrabaNombre_Click()

    PantaClaveNombre.Visible = False

End Sub

Private Sub WClaveNombre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WGrabaNombre = "N"
        WClaveNombre.Text = UCase(WClaveNombre.Text)
        If Trim(UCase(WClaveNombre.Text)) = "MERCEDEZ" Then
            WGrabaNombre = "S"
            PantaClaveNombre.Visible = False
            Call Mantenimiento_Click
                Else
            m$ = "Clave de Grabacion Invalida"
            a% = MsgBox(m$, 0, "Actualizacion de Nombres de P.T.")
            WClaveNombre.SetFocus
        End If
    End If

End Sub






Private Sub Command1_Click()

    Erase ZZVector
    ZZLugar = 0
    ZZPasa = 0

    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Precios"
    ZSql = ZSql + " Order by Terminado, Descripcion"
    spPrecios = ZSql
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    If rstPrecios.RecordCount > 0 Then
    
        With rstPrecios
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If Left$(UCase(rstPrecios!Terminado), 2) = "PT" Or Left$(UCase(rstPrecios!Terminado), 2) = "PE" Then
                
                        If ZZPasa = 0 Then
                            ZZCorte = UCase(rstPrecios!Terminado)
                            ZZCorteII = Trim(rstPrecios!Descripcion)
                            ZZRenglon = 0
                            ZZPasa = 1
                        End If
                    
                        If ZZCorte <> UCase(rstPrecios!Terminado) Or ZZCorteII <> Trim(rstPrecios!Descripcion) Then
                        
                            If ZZCorte <> UCase(rstPrecios!Terminado) Then
                                ZZRenglon = 0
                            End If
                            
                            ZZLugar = ZZLugar + 1
                            ZZRenglon = ZZRenglon + 1
                            ZZVector(ZZLugar, 1) = ZZCorte
                            ZZVector(ZZLugar, 2) = ZZCorteII
                            ZZVector(ZZLugar, 3) = ZZRenglon
                    
                            ZZCorte = UCase(rstPrecios!Terminado)
                            ZZCorteII = Trim(rstPrecios!Descripcion)
                            
                        End If
                        
                    End If
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstPrecios.Close
        
        ZZLugar = ZZLugar + 1
        ZZRenglon = ZZRenglon + 1
        ZZVector(ZZLugar, 1) = ZZCorte
        ZZVector(ZZLugar, 2) = ZZCorteII
        ZZVector(ZZLugar, 3) = ZZRenglon
        
    End If


    For ZZCiclo = 1 To ZZLugar
    
        ZZTerminado = ZZVector(ZZCiclo, 1)
        ZZDescripcion = ZZVector(ZZCiclo, 2)
        ZZRenglon = ZZVector(ZZCiclo, 3)
        
        Auxi = ZZRenglon
        Call Ceros(Auxi, 2)
        ZZClave = ZZTerminado + Auxi
        
        ZSql = ""
        ZSql = ZSql & "INSERT INTO PreciosII ("
        ZSql = ZSql & "Clave ,"
        ZSql = ZSql & "Terminado ,"
        ZSql = ZSql & "Renglon ,"
        ZSql = ZSql & "Nombre )"
        ZSql = ZSql & "Values ("
        ZSql = ZSql & "'" + ZZClave + "',"
        ZSql = ZSql & "'" + ZZTerminado + "',"
        ZSql = ZSql & "'" + ZZRenglon + "',"
        ZSql = ZSql & "'" + ZZDescripcion + "')"
    
        spPreciosII = ZSql
        Set rstPreciosII = db.OpenRecordset(spPreciosII, dbOpenSnapshot, dbSQLPassThrough)
        
    Next ZZCiclo


Stop
End Sub

