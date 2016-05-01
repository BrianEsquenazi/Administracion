VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgModif 
   AutoRedraw      =   -1  'True
   Caption         =   "Modificacion de Precios"
   ClientHeight    =   7275
   ClientLeft      =   1875
   ClientTop       =   840
   ClientWidth     =   8400
   LinkTopic       =   "Form2"
   ScaleHeight     =   7275
   ScaleWidth      =   8400
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
      Left            =   240
      TabIndex        =   42
      Top             =   120
      Width           =   4215
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
      Left            =   4560
      MaskColor       =   &H00808080&
      TabIndex        =   41
      Top             =   120
      Width           =   3735
   End
   Begin VB.Frame PantaDy 
      Height          =   6615
      Left            =   5400
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
      Begin VB.ListBox Pantalla1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   240
         TabIndex        =   32
         Top             =   4200
         Width           =   7335
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   29
         Text            =   " "
         Top             =   1200
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   28
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Porcentaje1 
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
         MaxLength       =   6
         TabIndex        =   27
         Text            =   " "
         Top             =   3240
         Width           =   1095
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
         Height          =   495
         Left            =   5520
         TabIndex        =   26
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Cancela1 
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
         Left            =   4080
         TabIndex        =   25
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton Acepta1 
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
         Left            =   4080
         TabIndex        =   24
         Top             =   2520
         Width           =   1335
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
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Visible         =   0   'False
         Width           =   7335
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
         Height          =   1740
         Left            =   2160
         TabIndex        =   22
         Top             =   4200
         Width           =   3615
      End
      Begin MSMask.MaskEdBox HastaArti 
         Height          =   375
         Left            =   2160
         TabIndex        =   30
         Top             =   2280
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
      Begin MSMask.MaskEdBox DesdeArti 
         Height          =   375
         Left            =   2160
         TabIndex        =   31
         Top             =   1800
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
      Begin VB.Label HastaDescri1 
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
         Left            =   3360
         TabIndex        =   40
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label DesdeDescri1 
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
         Left            =   3360
         TabIndex        =   39
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label9 
         Caption         =   "Porcentaje"
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
         TabIndex        =   38
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   37
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label11 
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
         Left            =   360
         TabIndex        =   36
         Top             =   1920
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
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "ACTUALIZACION DE PRECIOS DE MATERIAS PRUIMAS"
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
         Left            =   1080
         TabIndex        =   34
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label14 
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
         Height          =   375
         Left            =   360
         TabIndex        =   33
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.Frame PantaPt 
      Height          =   6735
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   4815
      Begin VB.TextBox Hastacliente 
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
         TabIndex        =   10
         Text            =   " "
         Top             =   1200
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   0
         Text            =   " "
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Porcentaje 
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
         MaxLength       =   6
         TabIndex        =   9
         Text            =   " "
         Top             =   3240
         Width           =   1095
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
         Left            =   5520
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
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
         Height          =   495
         Left            =   4080
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
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
         Height          =   495
         Left            =   4080
         TabIndex        =   6
         Top             =   2520
         Width           =   1335
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
         TabIndex        =   5
         Top             =   3840
         Visible         =   0   'False
         Width           =   7335
      End
      Begin VB.ListBox Pantalla 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2205
         Left            =   240
         TabIndex        =   4
         Top             =   4200
         Width           =   7335
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
         Height          =   1740
         Left            =   1920
         TabIndex        =   3
         Top             =   4200
         Width           =   3615
      End
      Begin MSMask.MaskEdBox HastaProd 
         Height          =   375
         Left            =   2160
         TabIndex        =   11
         Top             =   2400
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
      Begin MSMask.MaskEdBox DesdeProd 
         Height          =   375
         Left            =   2160
         TabIndex        =   12
         Top             =   1920
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
      Begin VB.Label HastaDescri 
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
         Left            =   3360
         TabIndex        =   20
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Label DesdeDescri 
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
         Left            =   3360
         TabIndex        =   19
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label5 
         Caption         =   "Porcentaje"
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
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   360
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label3 
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
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "ACTUALIZACION DE PRECIOS DE PRODUCTOS TERMINADOS"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label6 
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
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   -120
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   120
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "listsol.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Solicitudes de Conpras Realizadas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgModif"
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

Dim XParam As String

Dim Vector(10000)

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub

Sub Imprime_Descripcion()

    WCliente = DesdeCliente.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesdeDescri.Caption = rstCliente!Razon
            Else
        DesdeDescri.Caption = ""
    End If
    
    WCliente = Hastacliente.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        HastaDescri.Caption = rstCliente!Razon
            Else
        HastaDescri.Caption = ""
    End If
    
End Sub

Private Sub Acepta_Click()
    
    DesdeCliente.Text = UCase(DesdeCliente.Text)
    Hastacliente.Text = UCase(Hastacliente.Text)
    DesdeProd.Text = UCase(DesdeProd.Text)
    HastaProd.Text = UCase(HastaProd.Text)
                
    spPrecios = "ListaPrecios"
    Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    Erase Vector
    
    With rstPrecios
        .MoveFirst
        If .NoMatch = False Then
            Do
                If DesdeProd.Text <= rstPrecios!Terminado And HastaProd.Text >= rstPrecios!Terminado Then
                    If DesdeCliente.Text <= rstPrecios!Cliente And Hastacliente.Text >= rstPrecios!Cliente Then
                        Renglon = Renglon + 1
                        Vector(Renglon) = rstPrecios!Clave
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    For XX = 1 To 10000
        If Vector(XX) <> "" Then
            spPrecios = "ConsultaPrecios " + "'" + Vector(XX) + "'"
            Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
            If rstPrecios.RecordCount > 0 Then
                WPrecio = Str$(rstPrecios!Precio + (rstPrecios!Precio * Val(Porcentaje.Text) / 100))
                WClave = rstPrecios!Clave
                WCliente = rstPrecios!Cliente
                WTerminado = rstPrecios!Terminado
                WDescripcion = rstPrecios!Descripcion
                WDate = Date$
                rstPrecios.Close
                Rem by nan
                
                XParam = "'" + WClave + "','" + WPrecio + "','" + WDate + "'"
                Rem by nan
                spPrecios = "ModificaPrecios3 " + XParam
                Set rstPrecios = db.OpenRecordset(spPrecios, dbOpenSnapshot, dbSQLPassThrough)
              Rem by nan fin
                             
             Rem     Set rstPrecios = db.OpenRecordset("ModificaPrecios3 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    Next XX
    
    Call Cancela_click
    
End Sub

Private Sub Cancela_click()

    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeProd.Text = "  -     -   "
    HastaProd.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Opcion.Visible = False
    Pantalla.Visible = False
    
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -   -   "
    HastaArti.Text = "  -   -   "
    Porcentaje1.Text = ""
    DesdeDescri1.Caption = ""
    HastaDescri1.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False

    DesdeCliente.SetFocus
    PrgModif.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub DesdeCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente.Text = UCase(DesdeCliente.Text)
        Call Imprime_Descripcion
        Hastacliente.SetFocus
    End If
End Sub

Private Sub HastaCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hastacliente.Text = UCase(Hastacliente.Text)
        Call Imprime_Descripcion
        DesdeProd.SetFocus
    End If
End Sub

Private Sub DesdeProd_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeProd.Text = UCase(DesdeProd.Text)
        HastaProd.SetFocus
    End If
End Sub

Private Sub HastaProd_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaProd.Text = UCase(HastaProd.Text)
        Porcentaje.SetFocus
    End If
End Sub

Private Sub Porcentaje_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Porcentaje.Text = Pusing("###,###.##", Str$(Val(Porcentaje.Text)))
        DesdeCliente.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Form_Load()

    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeProd.Text = "  -     -   "
    HastaProd.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Opcion.Visible = False
    Pantalla.Visible = False
    
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -   -   "
    HastaArti.Text = "  -   -   "
    Porcentaje1.Text = ""
    DesdeDescri1.Caption = ""
    HastaDescri1.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False
    
    PantaDy.Visible = False
    PantaPt.Visible = True
    PantaPt.Height = 6615
    PantaPt.Left = 120
    PantaPt.Top = 480
    PantaPt.Width = 8175
    
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

    Opcion.Clear

    Opcion.AddItem "Clientes"
    Opcion.AddItem "Productos Terminados"

    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
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
            
        Case 1
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
            
            
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Text = ""
    Ayuda.Visible = True
    Ayuda.SetFocus

End Sub

Private Sub Ayuda_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 0
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
            
        Case 1
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
            WCliente = WIndice.List(Indice)
            DesdeCliente.Text = WCliente
            Hastacliente.Text = WCliente
            Call Imprime_Descripcion
            DesdeCliente.SetFocus
            
        Case 1
            Indice = Pantalla.ListIndex
            WTerminado = WIndice.List(Indice)
            DesdeProd.Text = WTerminado
            HastaProd.Text = WTerminado
            DesdeProd.SetFocus
            
        Case Else
    End Select
    
End Sub

Sub Imprime_Descripcion1()

    WCliente = DesdeCliente1.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        DesdeDescri1.Caption = rstCliente!Razon
            Else
        DesdeDescri1.Caption = ""
    End If
    
    WCliente = HastaCliente1.Text
    spCliente = "ConsultaCliente " + "'" + WCliente + "'"
    Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    If rstCliente.RecordCount > 0 Then
        HastaDescri1.Caption = rstCliente!Razon
            Else
        HastaDescri1.Caption = ""
    End If
    
End Sub

Private Sub Acepta1_Click()
    
    DesdeCliente1.Text = UCase(DesdeCliente1.Text)
    HastaCliente1.Text = UCase(HastaCliente1.Text)
    DesdeArti.Text = UCase(DesdeArti.Text)
    DesdeArti.Text = UCase(DesdeArti.Text)
                
    spPreciosMp = "ListaPreciosMP"
    Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
    
    Renglon = 0
    Erase Vector
    
    With rstPreciosMp
        .MoveFirst
        If .NoMatch = False Then
            Do
                If DesdeArti.Text <= rstPreciosMp!Articulo And HastaArti.Text >= rstPreciosMp!Articulo Then
                    If DesdeCliente1.Text <= rstPreciosMp!Cliente And HastaCliente1.Text >= rstPreciosMp!Cliente Then
                        Renglon = Renglon + 1
                        Vector(Renglon) = rstPreciosMp!Clave
                    End If
                End If
                .MoveNext
                If .EOF = True Then
                    Exit Do
                End If
            Loop
        End If
    End With
    
    For XX = 1 To 10000
        If Vector(XX) <> "" Then
            spPreciosMp = "ConsultaPreciosMp " + "'" + Vector(XX) + "'"
            Set rstPreciosMp = db.OpenRecordset(spPreciosMp, dbOpenSnapshot, dbSQLPassThrough)
            If rstPreciosMp.RecordCount > 0 Then
                WPrecio = Str$(rstPreciosMp!Precio + (rstPreciosMp!Precio * Val(Porcentaje1.Text) / 100))
                WClave = rstPreciosMp!Clave
                WCliente = rstPreciosMp!Cliente
                WArticulo = rstPreciosMp!Articulo
                WDate = Date$
                                     
                XParam = "'" + WClave + "','" + WPrecio + "','" + WDate + "'"
                Set rstPreciosMp = db.OpenRecordset("ModificaPreciosMp3 " + XParam, dbOpenSnapshot, dbSQLPassThrough)
            End If
        End If
    Next XX
    
    Call Cancela1_click
    
End Sub

Private Sub Cancela1_click()

    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeProd.Text = "  -     -   "
    DesdeProd.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False
    
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -   -   "
    HastaArti.Text = "  -   -   "
    Porcentaje1.Text = ""
    DesdeDescri1.Caption = ""
    HastaDescri1.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False

    DesdeCliente1.SetFocus
    PrgModif.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub DesdeCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeCliente1.Text = UCase(DesdeCliente1.Text)
        Call Imprime_Descripcion1
        HastaCliente1.SetFocus
    End If
End Sub

Private Sub HastaCliente1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaCliente1.Text = UCase(HastaCliente1.Text)
        Call Imprime_Descripcion1
        DesdeArti.SetFocus
    End If
End Sub

Private Sub DesdeArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DesdeArti.Text = UCase(DesdeArti.Text)
        HastaArti.SetFocus
    End If
End Sub

Private Sub HastaArti_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        HastaArti.Text = UCase(HastaArti.Text)
        Porcentaje1.SetFocus
    End If
End Sub

Private Sub Porcentaje1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Porcentaje1.Text = Pusing("###,###.##", Str$(Val(Porcentaje1.Text)))
        DesdeCliente1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Consulta1_Click()

    Opcion1.Visible = False
    Pantalla1.Visible = False

    Opcion1.Clear

    Opcion1.AddItem "Clientes"
    Opcion1.AddItem "Materias Primas"

    Opcion1.Visible = True
    
End Sub

Private Sub Opcion1_Click()

    Opcion1.Visible = False
     
    Dim IngresaItem As String

    Pantalla1.Clear
    WIndice.Clear

    XIndice = Opcion1.ListIndex
    
    Select Case XIndice
        Case 0
            spCliente = "ListaClienteConsulta"
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
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
            
        Case 1
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
            
        Case Else
    End Select
            
    Pantalla1.Visible = True
    Ayuda1.Text = ""
    Ayuda1.Visible = True
    Ayuda1.SetFocus

End Sub

Private Sub Ayuda1_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then

    Pantalla1.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda1.Text)
    
    Select Case XIndice
        Case 0
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
            
        Case 1
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
            WCliente = WIndice.List(Indice)
            DesdeCliente1.Text = WCliente
            HastaCliente1.Text = WCliente
            Call Imprime_Descripcion1
            DesdeCliente1.SetFocus
            
        Case 1
            Indice = Pantalla1.ListIndex
            WTerminado = WIndice.List(Indice)
            DesdeArti.Text = WTerminado
            HastaArti.Text = WTerminado
            DesdeArti.SetFocus
            
        Case Else
    End Select
    
End Sub

Private Sub BotonPt_Click()
    PantaDy.Visible = False
    PantaPt.Height = 6615
    PantaPt.Left = 120
    PantaPt.Top = 480
    PantaPt.Width = 8175
    PantaPt.Visible = True
    DesdeCliente.Text = ""
    Hastacliente.Text = ""
    DesdeProd.Text = "  -     -   "
    HastaProd.Text = "  -     -   "
    Porcentaje.Text = ""
    DesdeDescri.Caption = ""
    HastaDescri.Caption = ""
    Opcion.Visible = False
    Pantalla.Visible = False
    DesdeCliente.SetFocus
End Sub

Private Sub BotonDy_Click()
    PantaPt.Visible = False
    PantaDy.Visible = True
    PantaDy.Height = 6615
    PantaDy.Left = 120
    PantaDy.Top = 480
    PantaDy.Width = 8175
    DesdeCliente1.Text = ""
    HastaCliente1.Text = ""
    DesdeArti.Text = "  -   -   "
    HastaArti.Text = "  -   -   "
    Porcentaje1.Text = ""
    DesdeDescri1.Caption = ""
    HastaDescri1.Caption = ""
    Opcion1.Visible = False
    Pantalla1.Visible = False
    DesdeCliente1.SetFocus
End Sub

    
