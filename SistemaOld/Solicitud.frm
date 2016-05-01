VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgSolicitud 
   Caption         =   "Ingreso de Solicitud de Fondos"
   ClientHeight    =   8415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   11715
   Begin VB.Frame Frame3 
      Caption         =   "ESTADO"
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
      Height          =   2055
      Left            =   8640
      TabIndex        =   37
      Top             =   120
      Width           =   3015
      Begin VB.TextBox DescriPago2 
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
         MaxLength       =   15
         TabIndex        =   42
         Text            =   " "
         Top             =   1560
         Width           =   2775
      End
      Begin VB.TextBox DescriPago1 
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
         MaxLength       =   15
         TabIndex        =   41
         Text            =   " "
         Top             =   1200
         Width           =   2775
      End
      Begin VB.ComboBox Estado 
         Height          =   315
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox Orden 
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   38
         Text            =   " "
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Orden de Pago"
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
         TabIndex        =   40
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.TextBox Autorizado 
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
      TabIndex        =   36
      Top             =   1920
      Width           =   6255
   End
   Begin VB.TextBox Solicitante 
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
      TabIndex        =   34
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox Imputacion 
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
      TabIndex        =   28
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "OBSERVACIONES"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   8175
      Begin VB.TextBox Observaciones1 
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
         MaxLength       =   50
         TabIndex        =   27
         Top             =   240
         Width           =   7815
      End
      Begin VB.TextBox Observaciones2 
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
         MaxLength       =   50
         TabIndex        =   26
         Top             =   600
         Width           =   7815
      End
      Begin VB.TextBox Observaciones3 
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
         MaxLength       =   50
         TabIndex        =   25
         Top             =   960
         Width           =   7815
      End
   End
   Begin VB.TextBox Titulo 
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
      TabIndex        =   23
      Top             =   840
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "CONCEPTO"
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
      Height          =   2055
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   8175
      Begin VB.TextBox Descripcion5 
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
         MaxLength       =   50
         TabIndex        =   32
         Top             =   1680
         Width           =   7815
      End
      Begin VB.TextBox Descripcion4 
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
         MaxLength       =   50
         TabIndex        =   21
         Top             =   1320
         Width           =   7815
      End
      Begin VB.TextBox Descripcion3 
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
         MaxLength       =   50
         TabIndex        =   20
         Top             =   960
         Width           =   7815
      End
      Begin VB.TextBox Descripcion2 
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
         MaxLength       =   50
         TabIndex        =   19
         Top             =   600
         Width           =   7815
      End
      Begin VB.TextBox Descripcion1 
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
         MaxLength       =   50
         TabIndex        =   18
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.TextBox Solicitud 
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
      MaxLength       =   15
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox DesProveedor 
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
      MaxLength       =   50
      TabIndex        =   15
      Top             =   480
      Width           =   4815
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
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox Importe 
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
      Left            =   7200
      MaxLength       =   15
      TabIndex        =   12
      Text            =   " "
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Proveedor 
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
      TabIndex        =   10
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Top             =   120
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
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   6360
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
      Height          =   1980
      ItemData        =   "Solicitud.frx":0000
      Left            =   120
      List            =   "Solicitud.frx":0007
      TabIndex        =   5
      Top             =   6360
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "  Consulta      Datos       (F3)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9480
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "   Limpia        Pantalla    (F2)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9480
      TabIndex        =   3
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "   Fin de        Ingreso      (F10)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9480
      TabIndex        =   2
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "   Graba      (F1)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   9480
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin MSMask.MaskEdBox FechaPago 
      Height          =   285
      Left            =   5520
      TabIndex        =   30
      Top             =   1200
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
   Begin VB.Label Label8 
      Caption         =   "Autorizado por"
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
      Left            =   120
      TabIndex        =   35
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Solicitante"
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
      Left            =   120
      TabIndex        =   33
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha de Pago"
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
      TabIndex        =   31
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Imputacion Contable"
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
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
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
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Numero de Solicitud"
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
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Importe"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Beneficiario"
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
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha "
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
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   7
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstProveedor As Recordset
Dim spPrpveedor As String
Dim rstSolicitud As Recordset
Dim spSolicitud As String
Dim XParam As String
Dim XIndice As Integer
Dim WEstado As String

Private Sub cmdGraba_Click()

    If Val(Solicitud.Text) <> 0 Then
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdFechaPago = Right$(FechaPago.Text, 4) + Mid$(FechaPago.Text, 4, 2) + Left$(FechaPago.Text, 2)
        WEstado = Str$(Estado.ListIndex)
        
        XParam = "'" + WSolicitud + "','" _
                 + Fecha.Text + "','" _
                 + WOrdFecha + "','" _
                 + Importe.Text + "','" _
                 + Proveedor.Text + "','" _
                 + DesProveedor + "','" _
                 + WFechaord + "','" _
                 + Cantidad.Text + "','" _
                 + Cliente.Text + "','" _
                 + Razon.Text + "','" _
                 + DescriCliente.Text + "','" _
                 + Vendedor.Text + "','" _
                 + DesVendedor.Caption + "','" _
                 + Observaciones.Text + "'"
                 
        Set rstSolicitud = db.OpenRecordset("ModificaMuestraI " + XParam, dbOpenSnapshot, dbSQLPassThrough)
        Call cmdClose_Click

            Else

        WCodigo = 1
        spSolicitud = "ListaSolicitudNumero"
        Set rstSolicitud = db.OpenRecordset(spSolicitud, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolicitud.RecordCount > 0 Then
            With rstSolicitud
                .MoveLast
                WCodigo = rstSolicitud!Codigo + 1
            End With
            rstSolicitud.Close
        End If
    
        Solicitud.Text = Str$(WCodigo)
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WOrdFechaPago = Right$(FechaPago.Text, 4) + Mid$(FechaPago.Text, 4, 2) + Left$(FechaPago.Text, 2)
        WEstado = Str$(Estado.ListIndex)
        
        Sql1 = "INSERT INTO Solicitud ("
        Sql2 = "Solicitud ,"
        Sql3 = "Fecha ,"
        Sql4 = "OrdFecha ,"
        Sql5 = "Importe ,"
        Sql6 = "Proveedor ,"
        Sql7 = "DesProveedor ,"
        Sql8 = "Descripcion1 ,"
        Sql9 = "Descripcion2 ,"
        Sql10 = "Descripcion3 ,"
        Sql11 = "Descripcion4 ,"
        Sql12 = "Descripcion5 ,"
        Sql13 = "Imputacion ,"
        Sql14 = "FechaPago ,"
        Sql15 = "OrdFechaPago ,"
        Sql16 = "Observaciones1 ,"
        Sql17 = "Observaciones2 ,"
        Sql18 = "Observaciones3 ,"
        Sql19 = "Solicitante ,"
        Sql20 = "Autorizado ,"
        Sql21 = "Estado ,"
        Sql22 = "Orden ,"
        Sql23 = "DescriPago1 ,"
        Sql24 = "DescriPago2 ,"
        Sql25 = "Titulo )"
        Sql26 = "Values ("
        Sql27 = "'" + Solicitud.Text + "',"
        Sql28 = "'" + Fecha.Text + "',"
        Sql29 = "'" + WOrdFecha + "',"
        Sql30 = "'" + Importe.Text + "',"
        Sql31 = "'" + Proveedor.Text + "',"
        Sql32 = "'" + DesProveedor.Text + "',"
        Sql33 = "'" + Descripcion1.Text + "',"
        Sql34 = "'" + Descripcion2.Text + "',"
        Sql35 = "'" + Descripcion3.Text + "',"
        Sql36 = "'" + Descripcion4.Text + "',"
        Sql37 = "'" + Descripcion5.Text + "',"
        Sql38 = "'" + Imputacion.Text + "',"
        Sql39 = "'" + FechaPago.Text + "',"
        Sql40 = "'" + WOrdFechaPago + "',"
        Sql41 = "'" + Observaciones1.Text + "',"
        Sql42 = "'" + Observaciones2.Text + "',"
        Sql43 = "'" + Observaciones3.Text + "',"
        Sql44 = "'" + Solicitante.Text + "',"
        Sql45 = "'" + Autorizado.Text + "',"
        Sql46 = "'" + WEstado + "',"
        Sql47 = "'" + Orden.Text + "',"
        Sql48 = "'" + DescriPago1.Text + "',"
        Sql49 = "'" + DescriPago2.Text + "',"
        Sql50 = "'" + Titulo.Text + "')"
        
        spSolicitud = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 + _
                   Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18 + Sql19 + Sql20 + _
                   Sql21 + Sql22 + Sql23 + Sql24 + Sql25 + Sql26 + Sql27 + Sql28 + Sql29 + Sql30 + _
                   Sql31 + Sql32 + Sql33 + Sql34 + Sql35 + Sql36 + Sql37 + Sql38 + Sql39 + Sql40 + _
                   Sql41 + Sql42 + Sql43 + Sql44 + Sql45 + Sql46 + Sql47 + Sql48 + Sql49 + Sql50
        Set rstSolicitud = db.OpenRecordset(spSolicitud, dbOpenSnapshot, dbSQLPassThrough)
    
        Call CmdLimpiar_Click
        Solicitud.SetFocus
    End If
        
End Sub

Private Sub CmdLimpiar_Click()

    Producto.Text = "  -     -   "
    DesProducto.Caption = ""
    Articulo.Text = "  -   -   "
    DesArticulo.Caption = ""
    Ensayo.Text = ""
    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Cantidad.Text = ""
    Rem Cliente.Text = ""
    Rem Razon.Text = ""
    DescriCliente.Text = ""
    Rem Vendedor.Text = ""
    Rem DesVendedor.Caption = ""
    Rem Observaciones.Text = ""
    Producto.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgSolicitud.Hide
    Unload Me
    Menu.Show
End Sub

Sub Producto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Producto.Text <> "  -     -   " Then
        
            Producto.Text = UCase(Producto.Text)
            spTerminado = "ConsultaTerminado " + "'" + Producto.Text + "'"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesProducto.Caption = rstTerminado!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesProducto.Caption
                End If
                rstTerminado.Close
                Fecha.SetFocus
                    Else
                Producto.SetFocus
            End If
            
                Else
                
            Articulo.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Producto.Text = "  -     -   "
        DesProducto.Caption = ""
    End If
End Sub

Sub Articulo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Articulo.Text <> "  -   -   " Then
        
            Articulo.Text = UCase(Articulo.Text)
            spArticulo = "ConsultaArticulo " + "'" + Articulo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesArticulo.Caption = rstArticulo!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesArticulo.Caption
                End If
                rstArticulo.Close
                Fecha.SetFocus
                    Else
                Articulo.SetFocus
            End If
            
                Else
                
            Ensayo.SetFocus
            
        End If
    End If
    If KeyAscii = 27 Then
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
    End If
End Sub

Private Sub Ensayo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
    If KeyAscii = 27 Then
        Ensayo.Text = ""
    End If
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Cantidad.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Cantidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cliente.SetFocus
    End If
    If KeyAscii = 27 Then
        Cantidad.Text = ""
    End If
End Sub

Private Sub Cliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Cliente.Text <> "" And Cliente.Text <> Space$(6) Then
            Cliente.Text = UCase(Cliente.Text)
            spPrpveedor = "ConsultaCliente " + "'" + Cliente.Text + "'"
            Set RstProveedor = db.OpenRecordset(spPrpveedor, dbOpenSnapshot, dbSQLPassThrough)
            If RstProveedor.RecordCount > 0 Then
                Razon.Text = RstProveedor!Razon
                RstProveedor.Close
                DescriCliente.SetFocus
                    Else
                Cliente.SetFocus
            End If
                Else
            Razon.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
    End If
End Sub


Private Sub Razon_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescriCliente.SetFocus
    End If
    If KeyAscii = 27 Then
        Razon.Text = ""
    End If
End Sub

Private Sub DescriCliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Vendedor.SetFocus
    End If
    If KeyAscii = 27 Then
        DescriCliente.Text = ""
    End If
End Sub

Private Sub Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        spVendedor = "ConsultaVendedor " + "'" + Vendedor.Text + "'"
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            DesVendedor.Caption = rstVendedor!Nombre
            rstVendedor.Close
            Observaciones.SetFocus
                Else
            Vendedor.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Vendedor.Text = ""
        DesVendedor.Caption = ""
    End If
End Sub

Private Sub Observaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Producto.SetFocus
    End If
    If KeyAscii = 27 Then
        Observaciones.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Materias Primas"
    Opcion.AddItem "Clientes"
    Opcion.AddItem "Vendedores"
    
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
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstTerminado!Codigo, 2) = "PT" Then
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
            
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        Pantalla.AddItem IngresaItem
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
            
            
        Case 2
            spPrpveedors = "ListaClienteConsulta"
            Set rstProveedors = db.OpenRecordset(spPrpveedors, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedors.RecordCount > 0 Then
                With rstProveedors
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = rstProveedors!Cliente + " " + rstProveedors!Razon
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstProveedors!Cliente
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedors.Close
            End If
            
        Case 3
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstVendedor!Vendedor) + " " + rstVendedor!Nombre
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstVendedor!Vendedor
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
        
        Case Else
    End Select
            
    Pantalla.Visible = True
    Ayuda.Visible = True
    Ayuda.Text = ""
    Ayuda.SetFocus

End Sub

Private Sub Form_Load()

    Estado.Clear
    Estado.AddItem ""
    Estado.ListIndex = 0
    
    Solicitud.Text = ""
    Fecha.Text = "  /  /    "
    Importe.Text = ""
    Proveedor.Text = ""
    DesProveedor.Text = ""
    Descripcion1.Text = ""
    Descripcion2.Text = ""
    Descripcion3.Text = ""
    Descripcion4.Text = ""
    Descripcion5.Text = ""
    Imputacion.Text = ""
    FechaPago.Text = "  /  /    "
    Observaciones1.Text = ""
    Observaciones2.Text = ""
    Observaciones3.Text = ""
    Solicitante.Text = ""
    Autorizado.Text = ""
    Estado.ListIndex = 0
    Orden.Text = ""
    DescriPago1.Text = ""
    DescriPago2.Text = ""
    
End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Ayuda.Visible = False
    Select Case XIndice
        Case 0
            Indice = Pantalla.ListIndex
            ClavePro$ = WIndice.List(Indice)
            If Left$(ClavePro$, 2) = "PT" Then
                spTerminado = "ConsultaTerminado " + "'" + ClavePro$ + "'"
                Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
                If rstTerminado.RecordCount > 0 Then
                    Producto.Text = rstTerminado!Codigo
                    DesProducto.Caption = rstTerminado!Descripcion
                    If DescriCliente.Text = "" Then
                        DescriCliente.Text = DesProducto.Caption
                    End If
                    rstTerminado.Close
                        Else
                    Producto.Text = "  -     -   "
                    DesProducto.Caption = ""
                End If
                Producto.SetFocus
                    Else
                spArticulo = "ConsultaArticulo " + "'" + Left$(ClavePro$, 3) + Right$(ClavePro$, 7) + "'"
                Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
                If rstArticulo.RecordCount > 0 Then
                    Producto.Text = Left$(rstArticulo!Codigo, 3) + "00" + Right$(rstArticulo!Codigo, 7)
                    DesProducto.Caption = rstArticulo!Descripcion
                    If DescriCliente.Text = "" Then
                        DescriCliente.Text = DesProducto.Caption
                    End If
                    rstArticulo.Close
                        Else
                    Producto.Text = "  -   -   "
                    DesProducto.Caption = ""
                End If
                Producto.SetFocus
            End If
            
        Case 1
            Indice = Pantalla.ListIndex
            ClavePro$ = WIndice.List(Indice)
            spArticulo = "ConsultaArticulo " + "'" + Left$(ClavePro$, 3) + Right$(ClavePro$, 7) + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Articulo.Text = rstArticulo!Codigo
                DesArticulo.Caption = rstArticulo!Descripcion
                If DescriCliente.Text = "" Then
                    DescriCliente.Text = DesArticulo.Caption
                End If
                rstArticulo.Close
                    Else
                Articulo.Text = "  -   -   "
                DesArticulo.Caption = ""
            End If
            Articulo.SetFocus
            
        Case 2
            Indice = Pantalla.ListIndex
            Cliente.Text = WIndice.List(Indice)
            Call Cliente_KeyPress(13)
            
        Case 3
            Indice = Pantalla.ListIndex
            Vendedor.Text = WIndice.List(Indice)
            Call Vendedor_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    Opcion.Visible = False
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    WEspacios = Len(Ayuda.Text)
    
    Select Case XIndice
        Case 0
            spTerminado = "ListaTerminadoConsulta"
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        If Left$(rstTerminado!Codigo, 2) = "PT" Then
                            Da = Len(rstTerminado!Descripcion) - WEspacios
                            For Aaa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstTerminado!Descripcion, Aaa, WEspacios) Then
                                    IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstTerminado!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                        End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            End If
    
        Case 1
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
    
                With rstArticulo
                    .MoveFirst
                    Do
                        If .EOF = False Then
            
                            Da = Len(rstArticulo!Descripcion) - WEspacios
                            For Aaa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstArticulo!Descripcion, Aaa, WEspacios) Then
                                    IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstArticulo!Codigo
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                    
                                    Else
                        
                            Exit Do
                
                        End If
                    Loop
                End With
    
                rstArticulo.Close
            End If
            
        Case 2
            spPrpveedors = "ListaClienteConsulta"
            Set rstProveedors = db.OpenRecordset(spPrpveedors, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedors.RecordCount > 0 Then
                With rstProveedors
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstProveedors!Razon) - WEspacios
                            For Aaa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstProveedors!Razon, Aaa, WEspacios) Then
                                    IngresaItem = rstProveedors!Cliente + " " + rstProveedors!Razon
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstProveedors!Cliente
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstProveedors.Close
            End If
            
        Case 3
            spVendedor = "ListaVendedor"
            Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstVendedor.RecordCount > 0 Then
                With rstVendedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Da = Len(rstVendedor!Nombre) - WEspacios
                            For Aaa = 1 To Da
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstVendedor!Nombre, Aaa, WEspacios) Then
                                    IngresaItem = Str$(rstVendedor!Vendedor) + " " + rstVendedor!Nombre
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstVendedor!Vendedor
                                    WIndice.AddItem IngresaItem
                                    Exit For
                                End If
                            Next Aaa
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstVendedor.Close
            End If
    
        Case Else
    End Select
    
    End If

End Sub






Private Sub Producto_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Articulo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ensayo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cantidad_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Cliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Razon_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub DescriCliente_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Vendedor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Observaciones_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Ejecuta_Funcion()
    Select Case WFuncion
        Case 112
            Call cmdGraba_Click
        Case 113
            Call CmdLimpiar_Click
        Case 114
            Call Consulta_Click
        Case 121
            Call cmdClose_Click
        Case Else
    End Select
End Sub

Private Sub Form_Activate()
    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If
End Sub


