VERSION 5.00
Begin VB.Form prgcliagenda 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Clientes"
   ClientHeight    =   6510
   ClientLeft      =   945
   ClientTop       =   600
   ClientWidth     =   10020
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   10020
   Begin VB.CommandButton Cerrar 
      Caption         =   "Cerrar pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   52
      Top             =   5760
      Width           =   3975
   End
   Begin VB.TextBox DirEntrega 
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
      MaxLength       =   50
      TabIndex        =   49
      Text            =   " "
      Top             =   5160
      Width           =   5775
   End
   Begin VB.TextBox MInimo 
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
      Left            =   6240
      MaxLength       =   10
      TabIndex        =   48
      Text            =   " "
      Top             =   4800
      Width           =   1815
   End
   Begin VB.TextBox Limite 
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
      MaxLength       =   10
      TabIndex        =   47
      Text            =   " "
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox pago2 
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
      TabIndex        =   43
      Text            =   " "
      Top             =   4320
      Width           =   855
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
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   42
      Text            =   " "
      Top             =   3960
      Width           =   855
   End
   Begin VB.TextBox Horario 
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
      MaxLength       =   20
      TabIndex        =   41
      Text            =   " "
      Top             =   3600
      Width           =   2535
   End
   Begin VB.ComboBox Provincia 
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
      Left            =   2280
      TabIndex        =   37
      Text            =   " "
      Top             =   1440
      Width           =   3135
   End
   Begin VB.TextBox fax 
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
      MaxLength       =   40
      TabIndex        =   35
      Text            =   " "
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox email 
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
      MaxLength       =   40
      TabIndex        =   34
      Text            =   " "
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox Rubro 
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
      TabIndex        =   29
      Text            =   " "
      Top             =   3240
      Width           =   855
   End
   Begin VB.TextBox Vendedor 
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
      Left            =   6360
      MaxLength       =   4
      TabIndex        =   20
      Text            =   " "
      Top             =   2160
      Width           =   855
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
      Left            =   2280
      MaxLength       =   100
      TabIndex        =   19
      Text            =   " "
      Top             =   2880
      Width           =   3135
   End
   Begin VB.TextBox Contacto 
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
      MaxLength       =   50
      TabIndex        =   16
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Caption         =   "Condicion de Iva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      TabIndex        =   14
      Top             =   120
      Width           =   3375
      Begin VB.OptionButton Iva6 
         Caption         =   "No catalogado"
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
         Left            =   1680
         TabIndex        =   31
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Iva5 
         Caption         =   "Monotributo"
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
         TabIndex        =   27
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Iva4 
         Caption         =   "Exento"
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
         Left            =   1680
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton Iva3 
         Caption         =   "Cons. Final"
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
         Left            =   1680
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Iva2 
         Caption         =   "No Inscripto"
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
         TabIndex        =   23
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Iva1 
         Caption         =   "Inscripto"
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
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Cuit 
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
      MaxLength       =   13
      TabIndex        =   13
      Text            =   " "
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Telefono 
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
      Left            =   2280
      MaxLength       =   40
      TabIndex        =   12
      Text            =   " "
      Top             =   2160
      Width           =   3135
   End
   Begin VB.TextBox Postal 
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
      MaxLength       =   10
      TabIndex        =   11
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox Localidad 
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
      MaxLength       =   50
      TabIndex        =   10
      Text            =   " "
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Direccion 
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
      MaxLength       =   50
      TabIndex        =   5
      Text            =   " "
      Top             =   720
      Width           =   3135
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
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Razon 
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
      MaxLength       =   50
      TabIndex        =   3
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Despago2 
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
      Left            =   3360
      TabIndex        =   51
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Despago1 
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
      Left            =   3360
      TabIndex        =   50
      Top             =   3960
      Width           =   3255
   End
   Begin VB.Label Label16 
      Caption         =   "Direccion de Entrega"
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
      TabIndex        =   46
      Top             =   5160
      Width           =   2055
   End
   Begin VB.Label Label15 
      Caption         =   "Minimo a Facturar"
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
      Left            =   4200
      TabIndex        =   45
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Limite de Credito"
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
      TabIndex        =   44
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "Condicion de Proyeccion"
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
      Left            =   120
      TabIndex        =   40
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label10 
      Caption         =   "Horario"
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
      TabIndex        =   38
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Provincia"
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
      TabIndex        =   36
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label20 
      Caption         =   "Fax"
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
      Left            =   5520
      TabIndex        =   33
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "E-Mail"
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
      Left            =   5520
      TabIndex        =   32
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label DesRubro 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   3240
      Width           =   3375
   End
   Begin VB.Label Label17 
      Caption         =   "Rubro"
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
      TabIndex        =   28
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label19 
      Caption         =   " "
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label DesVendedor 
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
      Left            =   7320
      TabIndex        =   21
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
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
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Provi 
      Caption         =   "Contacto"
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
      TabIndex        =   15
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Cuit"
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
      Left            =   3600
      TabIndex        =   9
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "Telefono"
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
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo Postal"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Poblaci 
      Caption         =   "Localidad"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Direccion"
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblLabels 
      Caption         =   "Razon Social"
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
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo de Cliente"
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
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1815
   End
End
Attribute VB_Name = "prgcliagenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstRubro As Recordset
Dim spRubro As String
Dim rstPago As Recordset
Dim spPago As String
Dim XParam As String

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Sub Imprime_Descripcion()


    Rem lee rubro

    WRubro = Rubro.Text
    spRubro = "ConsultaRubro " + "'" + Rubro.Text + "'"
    Set rstRubro = db.OpenRecordset(spRubro, dbOpenSnapshot, dbSQLPassThrough)
    If rstRubro.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
        DesRubro.Caption = rstRubro!Nombre
            Else
        DesRubro.Caption = ""
    End If
    
    Rem lee vendedor
    WVendedor = vendedor.Text
    spVendedor = "ConsultaVendedor " + "'" + vendedor.Text + "'"
    Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
    If rstVendedor.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
        DesVendedor.Caption = rstVendedor!Nombre
            Else
        DesVendedor.Caption = ""
    End If

    
    Rem lee condicion de pago 1
    
    WPago1 = Pago1.Text
    spPago = "ConsultaPago " + "'" + Pago1.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
        Despago1.Caption = rstPago!Nombre
            Else
        Despago1.Caption = ""
    End If

    Rem lee condicion de pago 2
    
    WPago2 = Pago2.Text
    spPago = "ConsultaPago " + "'" + Pago2.Text + "'"
    Set rstPago = db.OpenRecordset(spPago, dbOpenSnapshot, dbSQLPassThrough)
    If rstPago.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
        Despago2.Caption = rstPago!Nombre
            Else
        Despago2.Caption = ""
    End If
    
End Sub

Sub Format_datos()
    Limite.Text = Pusing("###,###.##", Limite.Text)
    MInimo.Text = Pusing("###,###.##", MInimo.Text)
End Sub

Sub Imprime_Datos()

    spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        WPasa = "S"
            Else
        WPasa = "N"
    End If
                
    If WPasa = "S" Then
            Cliente.Text = rstCliente!Cliente
            Razon.Text = rstCliente!Razon
            Direccion.Text = rstCliente!Direccion
            Localidad.Text = rstCliente!Localidad
            Postal.Text = rstCliente!Postal
            Telefono.Text = rstCliente!Telefono
            Contacto.Text = rstCliente!Contacto
            Observaciones.Text = rstCliente!Observaciones
            Cuit.Text = rstCliente!Cuit
            vendedor.Text = rstCliente!vendedor
            email.Text = rstCliente!email
            fax.Text = rstCliente!fax
            Rubro.Text = rstCliente!Rubro
            Horario.Text = rstCliente!Horario
            Pago1.Text = rstCliente!Pago1
            Pago2.Text = rstCliente!Pago2
            Limite.Text = rstCliente!Limite
            MInimo.Text = rstCliente!MInimo
            DirEntrega.Text = rstCliente!DirEntrega
            Iva1.Value = False
            Iva2.Value = False
            Iva3.Value = False
            Iva4.Value = False
            Iva5.Value = False
            Iva6.Value = False
            Provincia.ListIndex = rstCliente!Provincia
            Select Case Val(rstCliente!Iva)
                Case 1
                    Iva1.Value = True
                Case 2
                    Iva2.Value = True
                Case 3
                    Iva3.Value = True
                Case 4
                    Iva4.Value = True
                Case 5
                    Iva5.Value = True
                Case 6
                    Iva6.Value = True
                Case Else
            End Select
            Call Format_datos
            Call Imprime_Descripcion
    End If
End Sub


Private Sub CmdLimpiar_Click()
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    vendedor.Text = ""
    DesVendedor.Caption = ""
    email.Text = ""
    fax.Text = ""
    Rubro.Text = ""
    DesRubro.Caption = ""
    Horario.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    Limite.Text = ""
    MInimo.Text = ""
    DirEntrega.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    DesRubro.Caption = ""
    Despago1.Caption = ""
    Despago2.Caption = ""
    Provincia.ListIndex = 25
    Cliente.SetFocus
End Sub

Private Sub Cerrar_Click()
    Call CmdLimpiar_Click
    Cliente.SetFocus
    prgcliagenda.Hide
    Unload Me
    PrgAltaAgenda.Show
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.Text = UCase(Cliente.Text)
        If Cliente.Text <> "" Then
            If Len(Cliente.Text) < 6 Then
                m$ = "El codigo de cliente debe tener 6 digitos"
                a% = MsgBox(m$, 0, "Archivo de Cliente")
                Cliente.SetFocus
                    Else
                spCliente = "ConsultaCliente " + "'" + Cliente.Text + "'"
                Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                If rstCliente.RecordCount > 0 Then
                    WPasa = "S"
                        Else
                    WPasa = "N"
                End If
                If WPasa = "S" Then
                        Cliente.Text = rstCliente!Cliente
                        Call Imprime_Datos
                            Else
                        CmdLimpiar_Click
                        Cliente.Text = Claveven$
                End If
                Razon.SetFocus
            End If
        End If
    End If
End Sub

Sub Form_Load()
    Cliente.Text = ""
    Razon.Text = ""
    Direccion.Text = ""
    Localidad.Text = ""
    Postal.Text = ""
    Telefono.Text = ""
    Contacto.Text = ""
    Observaciones.Text = ""
    Cuit.Text = ""
    vendedor.Text = ""
    DesVendedor.Caption = ""
    email.Text = ""
    fax.Text = ""
    Rubro.Text = ""
    DesRubro.Caption = ""
    Horario.Text = ""
    Pago1.Text = ""
    Pago2.Text = ""
    Limite.Text = ""
    MInimo.Text = ""
    DirEntrega.Text = ""
    Iva1.Value = True
    Iva2.Value = False
    Iva3.Value = False
    Iva4.Value = False
    Iva5.Value = False
    Iva6.Value = False
    Despago1.Caption = ""
    Despago2.Caption = ""
    
    Provincia.Clear
    
    Provincia.AddItem "Capital Federal"
    Provincia.AddItem "Buenos Aires"
    Provincia.AddItem "Catamarca"
    Provincia.AddItem "Cordoba"
    Provincia.AddItem "Corrientes"
    Provincia.AddItem "Chaco"
    Provincia.AddItem "Chubut"
    Provincia.AddItem "Entre Rios"
    Provincia.AddItem "Formosa"
    Provincia.AddItem "Jujuy"
    Provincia.AddItem "La Pampa"
    Provincia.AddItem "La Rioja"
    Provincia.AddItem "Mendoza"
    Provincia.AddItem "Misiones"
    Provincia.AddItem "Neuquen"
    Provincia.AddItem "Rio Negro"
    Provincia.AddItem "Salta"
    Provincia.AddItem "San Juan"
    Provincia.AddItem "San Luis"
    Provincia.AddItem "Santa Cruz"
    Provincia.AddItem "Santa Fe"
    Provincia.AddItem "Santiago del Estero"
    Provincia.AddItem "Tucuman"
    Provincia.AddItem "Tierra del Fuego"
    Provincia.AddItem "Exterior"
    Provincia.AddItem ""
    
    Cliente.Text = PCliente
    Call Imprime_Datos
    
End Sub
