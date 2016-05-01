VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgAtrasoEntrega 
   Caption         =   "Ingreso de Aviso de No entrega"
   ClientHeight    =   8415
   ClientLeft      =   1665
   ClientTop       =   405
   ClientWidth     =   8580
   LinkTopic       =   "Form2"
   ScaleHeight     =   8415
   ScaleWidth      =   8580
   Begin VB.TextBox Emisor 
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
      Left            =   5280
      MaxLength       =   50
      TabIndex        =   30
      Text            =   " "
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Solicitud 
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
      MaxLength       =   65
      TabIndex        =   27
      Text            =   " "
      Top             =   1920
      Width           =   1335
   End
   Begin VB.ComboBox Concepto 
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
      TabIndex        =   26
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox Pedido 
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
      MaxLength       =   65
      TabIndex        =   0
      Text            =   " "
      Top             =   480
      Width           =   1335
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
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   16
      Text            =   " "
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Problema 
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
      TabIndex        =   15
      Text            =   " "
      Top             =   1560
      Width           =   6015
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   2280
      TabIndex        =   14
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
   Begin MSMask.MaskEdBox Terminado 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   1080
      TabIndex        =   10
      Top             =   5040
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   3000
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
      Height          =   3660
      ItemData        =   "atrasoentrega.frx":0000
      Left            =   120
      List            =   "atrasoentrega.frx":0007
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "   Consulta       Datos           (F3)"
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
      Left            =   4560
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "    Limpia         Pantalla          (F2)"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "    Fin de          Ingreso         (F10)"
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
      Left            =   6360
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGraba 
      Caption         =   "    Graba           (F1)"
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
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
   Begin MSMask.MaskEdBox Articulo 
      Height          =   285
      Left            =   2280
      TabIndex        =   19
      Top             =   2280
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox FechaEntrega 
      Height          =   285
      Left            =   2280
      TabIndex        =   22
      Top             =   2640
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
   Begin VB.Label Label10 
      Caption         =   "Quien lo emite"
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
      Left            =   3720
      TabIndex        =   29
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Solic. Materia Prima"
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
      TabIndex        =   28
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo de Retraso"
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
      TabIndex        =   25
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label DesCliente 
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
      Left            =   3720
      TabIndex        =   24
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Entrega Estimada"
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
      TabIndex        =   23
      Top             =   2640
      Width           =   2415
   End
   Begin VB.Label DesArticulo 
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
      Left            =   3720
      TabIndex        =   21
      Top             =   2280
      Width           =   4575
   End
   Begin VB.Label Label8 
      Caption         =   "Materia  Prima"
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
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Numero de Pedido"
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
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Fecha del Aviso"
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
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Problema"
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
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label DesTerminado 
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1200
      Width           =   4575
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   8
      Top             =   3360
      Width           =   375
   End
End
Attribute VB_Name = "PrgAtrasoEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim rstCliente As Recordset
Dim spCliente As String
Dim rstAtraso As Recordset
Dim spAtraso As String
Dim rstVendedor As Recordset
Dim spVendedor As String
Dim rstSolic As Recordset
Dim spSolic As String
Dim XParam As String
Dim EmpresaActual As String
Dim XIndice As Integer
Dim EmailAddress As String
Dim WEmail(100) As String
Dim ret As Long
Dim sTo As String
Dim sCC As String
Dim sBCC As String
Dim sSubject As String
Dim sBody As String
Dim MSubject As String
Dim MBody As String
Dim AllPath As String

Dim ZZDesCliente As String

Private Sub cmdGraba_Click()
Existe = 0
    Rem On Error GoTo WError
    
    If Concepto.ListIndex <= 0 Then
        m$ = "Se debe informar el concpeto del atraso"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    
    If Concepto.ListIndex = 1 Or Concepto.ListIndex = 2 Then
        Sql1 = "Select *"
        Sql2 = " FROM Articulo"
        Sql3 = " Where Articulo.Codigo = " + "'" + Articulo.Text + "'"
        spArticulo = Sql1 + Sql2 + Sql3
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            rstArticulo.Close
                Else
            m$ = "Se debe informar el codigo de materia prima que esta en falta"
            a% = MsgBox(m$, 0, "Ingreso de Aviso de Atraso de Entragado")
            Exit Sub
        End If
    End If
    
    If Val(Pedido.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        Sql4 = " Order by Clave"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount = 0 Then
            m$ = "Numero de Pedido Incorrecto"
            a% = MsgBox(m$, 0, "Aviso de Atraso")
            Exit Sub
                Else
            rstPedido.Close
        End If
    End If
    
    If Val(Solicitud.Text) <> 0 Then
        Sql1 = "Select *"
        Sql2 = " FROM Solic"
        Sql3 = " Where Solic.Solicitud = " + "'" + Solicitud.Text + "'"
        spSolic = Sql1 + Sql2 + Sql3
        Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
        If rstSolic.RecordCount = 0 Then
            m$ = "Numero de Pedido Incorrecto"
            a% = MsgBox(m$, 0, "Aviso de Atraso")
            Exit Sub
                Else
            rstSolic.Close
        End If
    End If
    
    Rem by nan
    If Cliente.Text = "" Then
        m$ = "Se debe ingresar Cliente"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    
    If Problema.Text = "" Then
        m$ = "Se debe ingresar Problema"
        a% = MsgBox(m$, 0, "Aviso de Atraso")
        Exit Sub
    End If
    Rem fin nan
    
    If Val(WAtraso) <> 0 Then
    
        WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        WFechaEntregaord = Right$(FechaEntrega.Text, 4) + Mid$(FechaEntrega.Text, 4, 2) + Left$(FechaEntrega.Text, 2)
        
        Sql1 = "UPDATE Atraso SET "
        Sql2 = "Numero = " + "'" + WAtraso + "',"
        Sql3 = "Fecha = " + "'" + Fecha.Text + "',"
        Sql4 = "OrdFecha = " + "'" + WFechaOrd + "',"
        Sql5 = "Pedido = " + "'" + Pedido.Text + "',"
        Sql6 = "Cliente = " + "'" + Cliente.Text + "',"
        Sql8 = "Terminado = " + "'" + Terminado.Text + "',"
        Sql9 = "Problema = " + "'" + Problema.Text + "',"
        Sql10 = "Articulo = " + "'" + Articulo.Text + "',"
        Sql11 = "FechaEntrega = " + "'" + FechaEntrega.Text + "',"
        Sql12 = "OrdFechaEntrega = " + "'" + WOrdFechaEntrega + "',"
        Sql13 = "DesCliente = " + "'" + DesCliente.Caption + "',"
        Sql14 = "DesTerminado = " + "'" + DesTerminado.Caption + "',"
        Sql15 = "DesArticulo = " + "'" + DesArticulo.Caption + "',"
        Sql16 = "Concepto = " + "'" + Str$(Concepto.ListIndex) + "',"
        Sql17 = "Solicitud = " + "'" + Solicitud.Text + "'"
        Sql18 = " Where Numero = " + "'" + WAtraso + "'"
                     
        spAtraso = Sql1 + Sql2 + Sql3 + Sql4 + Sql5 + Sql6 + Sql7 + Sql8 + Sql9 + Sql10 _
                 + Sql11 + Sql12 + Sql13 + Sql14 + Sql15 + Sql16 + Sql17 + Sql18
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
     Rem   Call cmdClose_Click
          Existe = 1
         End If
            
           Rem by nan quito else
            
         Rem   Else
             
             Rem fin by nan
             If Existe = 0 Then
                  WAtraso = "1"
        
                    Sql1 = "Select Max(Numero) as [NumeroMayor]"
                    Sql2 = " FROM Atraso"
                   spAtraso = Sql1 + Sql2
                  Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
                   If rstAtraso.RecordCount > 0 Then
                      WAtraso = Str$(rstAtraso!Numeromayor + 1)
                     rstAtraso.Close
                    End If
        
                 ZZVersionPedido = ""
                 Sql1 = "Select *"
                 Sql2 = " FROM Pedido"
                 Sql3 = " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
                 Sql4 = " Order by Clave"
                 spPedido = Sql1 + Sql2 + Sql3 + Sql4
                 Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
                If rstPedido.RecordCount > 0 Then
                 ZZVersionPedido = Str$(rstPedido!Version)
                 rstPedido.Close
                End If
   
        
       
             WFechaOrd = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
             WFechaEntregaord = Right$(FechaEntrega.Text, 4) + Mid$(FechaEntrega.Text, 4, 2) + Left$(FechaEntrega.Text, 2)
        
              ZSql = ""
              ZSql = ZSql + "INSERT INTO Atraso ("
              ZSql = ZSql + "Numero ,"
              ZSql = ZSql + "Fecha ,"
              ZSql = ZSql + "OrdFecha ,"
              ZSql = ZSql + "Pedido ,"
              ZSql = ZSql + "Cliente ,"
              ZSql = ZSql + "Terminado ,"
              ZSql = ZSql + "Problema ,"
              ZSql = ZSql + "Articulo ,"
              ZSql = ZSql + "FechaEntrega ,"
              ZSql = ZSql + "OrdFechaEntrega ,"
              ZSql = ZSql + "DesCliente ,"
              ZSql = ZSql + "DesTerminado ,"
              ZSql = ZSql + "DesArticulo ,"
              ZSql = ZSql + "Concepto ,"
              ZSql = ZSql + "Solicitud ,"
              ZSql = ZSql + "Origen ,"
              ZSql = ZSql + "VersionPedido)"
              ZSql = ZSql + "Values ("
              ZSql = ZSql + "'" + WAtraso + "',"
              ZSql = ZSql + "'" + Fecha.Text + "',"
              ZSql = ZSql + "'" + WOrdFecha + "',"
              ZSql = ZSql + "'" + Pedido.Text + "',"
              ZSql = ZSql + "'" + Cliente.Text + "',"
              ZSql = ZSql + "'" + Terminado.Text + "',"
              ZSql = ZSql + "'" + Problema.Text + "',"
              ZSql = ZSql + "'" + Articulo.Text + "',"
              ZSql = ZSql + "'" + FechaEntrega.Text + "',"
              ZSql = ZSql + "'" + WOrdFechaEntrega + "',"
              ZSql = ZSql + "'" + DesCliente.Caption + "',"
              ZSql = ZSql + "'" + DesTerminado.Caption + "',"
              ZSql = ZSql + "'" + DesArticulo.Caption + "',"
              ZSql = ZSql + "'" + Str$(Concepto.ListIndex) + "',"
              ZSql = ZSql + "'" + Solicitud.Text + "',"
              ZSql = ZSql + "'" + "0" + "',"
              ZSql = ZSql + "'" + ZZVersionPedido + "')"
        
             spAtraso = ZSql
           Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
     End If
      
      
      Rem fin by nan si no existe
        Rem Busca el Cliente
        
        WCliente = ""
        ZZVersionPedido = ""
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        Sql4 = " Order by Clave"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            ZZVersionPedido = Str$(rstPedido!Version)
            WCliente = rstPedido!Cliente
            rstPedido.Close
        End If
        
        Rem Busca el Vendedor
        
        WVendedor = ""
        Sql1 = "Select *"
        Sql2 = " FROM Cliente"
        Sql3 = " Where Cliente.Cliente = " + "'" + WCliente + "'"
        spCliente = Sql1 + Sql2 + Sql3
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            WVendedor = Str$(rstCliente!vendedor)
            rstCliente.Close
        End If
        
        ZZDesCliente = Trim(DesCliente.Caption)
        For Ciclo = 1 To Len(ZZDesCliente)
            If Mid$(ZZDesCliente, Ciclo, 1) = "&" Then
                ZZDesCliente = Mid$(ZZDesCliente, 1, Ciclo - 1) + " " + Mid$(ZZDesCliente, Ciclo + 1, 100)
                ZZDesCliente = Trim(ZZDesCliente)
            End If
        Next Ciclo
        
        Rem Busca el email del Vendedor
        
        WEmail(7) = ""
        WEmail(8) = ""
        XEmail = ""
        Sql1 = "Select *"
        Sql2 = " FROM Vendedor"
        Sql3 = " Where Vendedor.Vendedor = " + "'" + WVendedor + "'"
        spVendedor = Sql1 + Sql2 + Sql3
        Set rstVendedor = db.OpenRecordset(spVendedor, dbOpenSnapshot, dbSQLPassThrough)
        If rstVendedor.RecordCount > 0 Then
            XEmail = Trim(rstVendedor!Email1)
            WEmail(8) = Trim(rstVendedor!Email2)
            rstVendedor.Close
        End If
        
        Rem by nan
        If Wempresa = "0008" Then
            WEmail(1) = "Aviso Pellital"
            WEmail(2) = ""
            WEmail(3) = ""
            WEmail(4) = ""
            WEmail(5) = ""
            WEmail(6) = ""
               Else
            WEmail(1) = "Aviso"
            WEmail(2) = ""
            WEmail(3) = ""
            WEmail(4) = ""
            WEmail(5) = ""
            WEmail(6) = ""
        End If
        
        XTipoPro = ""
        If Val(Wempresa) = 1 Then
            XCodigo = Val(Mid$(Terminado.Text, 4, 5))
            If Left$(Terminado.Text, 2) = "DY" Or Left$(Terminado.Text, 2) = "DW" Then
                XTipoPro = "CO"
                    Else
                If XCodigo >= 0 And XCodigo <= 999 Then
                    XTipoPro = "CO"
                        Else
                    If XCodigo >= 11000 And XCodigo <= 11999 Then
                        XTipoPro = "CO"
                    End If
                End If
            End If
        End If
        
        If Val(Wempresa) = 1 Then
            If XTipoPro = "CO" Then
                WEmail(1) = "Aviso;Colorante"
            End If
        End If
        
        For Ciclo = 1 To 8
        
            If WEmail(Ciclo) <> "" Then
            
                sTo = WEmail(Ciclo)
                
                If Ciclo <> 99 Then
        
                    sCC = XEmail
                    sBCC = ""
                    sSubject = "AVISO DE NO ENTREGA DE P.T."
                    If Articulo.Text = "  -   -   " Then
                        sBody = "Pedido:" + Pedido.Text + " - " + _
                                "Cliente:" + RTrim(DesCliente.Caption) + " - " + _
                                "Producto:" + Terminado.Text + " " + RTrim(DesTerminado.Caption) + " - " + _
                                "Problema:" + RTrim(Problema.Text) + " - " + _
                                "Fecha Estimada de Entrega:" + FechaEntrega.Text
                                        Else
                        sBody = "Pedido:" + Pedido.Text + " - " + _
                                "Cliente:" + RTrim(DesCliente.Caption) + " - " + _
                                "Producto:" + Terminado.Text + " " + RTrim(DesTerminado.Caption) + " - " + _
                                "Problema:" + RTrim(Problema.Text) + " - " + _
                                "M.P.:" + Articulo.Text + " " + RTrim(DesArticulo.Caption) + " - " + _
                                "Fecha Estimada de Entrega:" + FechaEntrega.Text
                    End If
    
                    Rem by nan  22-5-2014 RUTINA OUTLOOK
        
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
                       
                    Rem by nan
                    Rem    ret = Shell("Start.exe " _
                    REM        & "mailto:" & """" & sTo & """" _
                    REM        & "?Subject=" & """" & sSubject & """" _
                    REM        & "&cc=" & """" & sCC & """" _
                    REM        & "&bcc=" & """" & sBCC & """" _
                    REM & "&Body=" & """" & sBody & """" _
                    REM & "&File=" & """" & "c:\autoexec.bat" & """" _
                    REM  , 0)
                    
                        Else
                            
                    sCC = ""
                    sBCC = ""
                    sSubject = "NO ENTREGA DE P.T.:" + Terminado.Text + " - "
                    sBody = RTrim(DesCliente.Caption)
            
                    Rem  PARA OUTLOOK EXPRESS
                    Rem    ret = Shell("Start.exe " _
                    REM   & "mailto:" & """" & sTo & """" _
                    REM   & "?Subject=" & """" & sSubject & """" _
                    REM   & "&cc=" & """" & sCC & """" _
                    REM & "&bcc=" & """" & sBCC & """" _
                    REM & "&Body=" & """" & sBody & """" _
                    REM & "&File=" & """" & "c:\autoexec.bat" & """" _
                    REM , 0)
                        
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
                
                Inicio = Timer
                Do
                    Final = Timer
                    Dife = Final - Inicio
                    If Dife > 3 Then
                        Exit Do
                    End If
                Loop
                
            End If
            
        Next Ciclo
        
        WAtraso = ""
        Rem Call CmdLimpiar_Click
        Fecha.SetFocus
        
 Rem   End If

    Exit Sub

WError:
    Resume Next
        
End Sub

Private Sub CmdLimpiar_Click()

    Concepto.ListIndex = 0

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Problema.Text = ""
    Articulo.Text = "  -   -   "
    FechaEntrega.Text = "  /  /    "
    DesCliente.Caption = ""
    DesTerminado.Caption = ""
    DesArticulo.Caption = ""
    Fecha.SetFocus
    
End Sub

Private Sub cmdClose_Click()
    PrgAtrasoEntrega.Hide
    Unload Me
    PrgMuestraAtraso.Show
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Pedido.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Pedido_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        Sql4 = " Order by Clave"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            Cliente.Text = rstPedido!Cliente
            rstPedido.Close
            Sql1 = "Select *"
            Sql2 = " FROM Cliente"
            Sql3 = " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
            spCliente = Sql1 + Sql2 + Sql3
            Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstCliente.RecordCount > 0 Then
                DesCliente.Caption = rstCliente!razon
                rstCliente.Close
            End If
            Terminado.SetFocus
                Else
            T$ = "Ingreso de Aviso de No Entrega"
            m$ = "Pedido Inexistente"
            a% = MsgBox(m$, 0, T$)
            Pedido.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Pedido.Text = ""
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Cliente_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Sql1 = "Select *"
        Sql2 = " FROM CLiente"
        Sql3 = " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        Sql4 = " Order by Cliente"
        spCliente = Sql1 + Sql2 + Sql3 + Sql4
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!razon
            rstCliente.Close
            Terminado.SetFocus
                Else
            m$ = "Cliente Inexistente"
            a% = MsgBox(m$, 0, "Ingreso de Aviso de Atraso de Entragado")
            Cliente.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Cliente.Text = ""
        DesCliente.Caption = ""
    End If
End Sub

Sub Terminado_Keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        Terminado.Text = UCase(Terminado.Text)
        
        Sql1 = "Select *"
        Sql2 = " FROM Pedido"
        Sql3 = " Where Pedido.Pedido = " + "'" + Pedido.Text + "'"
        Sql4 = " and Pedido.Terminado = " + "'" + Terminado.Text + "'"
        spPedido = Sql1 + Sql2 + Sql3 + Sql4
        Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)
        If rstPedido.RecordCount > 0 Then
            rstPedido.Close
                Else
            T$ = "Ingreso de Aviso de No Entrega"
            m$ = "No se encuentra cargado el Producto en el pedido especificado"
            a% = MsgBox(m$, 0, T$)
            Exit Sub
        End If
        
        If Left$(Terminado.Text, 2) <> "DY" And Left$(Terminado.Text, 2) <> "DW" Then
            
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesTerminado.Caption = rstTerminado!Descripcion
                rstTerminado.Close
                    Else
                m$ = "Producto Terminado Inexistente"
                a% = MsgBox(m$, 0, "Ingreso de Aviso de Atraso de Entragado")
                Exit Sub
            End If
            
            Problema.SetFocus
            
                Else
                
            WArticulo = Left$(Terminado.Text, 3) + Right$(Terminado.Text, 7)
                
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesTerminado.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                m$ = "Producto Terminado Inexistente"
                a% = MsgBox(m$, 0, "Ingreso de Aviso de Atraso de Entragado")
                Exit Sub
            End If
            
            Problema.SetFocus
            
        End If
        
    End If
    
    If KeyAscii = 27 Then
        Terminado.Text = "  -     -   "
        DesTerminado.Caption = ""
    End If
    
End Sub

Private Sub Problema_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Solicitud.SetFocus
    End If
    If KeyAscii = 27 Then
        Problema.Text = ""
    End If
End Sub

Sub Solicitud_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Solicitud.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM Solic"
            Sql3 = " Where Solic.Solicitud = " + "'" + Solicitud.Text + "'"
            spSolic = Sql1 + Sql2 + Sql3
            Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
            If rstSolic.RecordCount > 0 Then
                rstSolic.Close
                Articulo.SetFocus
                    Else
                T$ = "Ingreso de Aviso de No Entrega"
                m$ = "Numero de Solicitud Inexistente"
                a% = MsgBox(m$, 0, T$)
                Solicitud.SetFocus
            End If
                Else
            Articulo.SetFocus
        End If
    End If
    
    If KeyAscii = 27 Then
        Solicitud.Text = ""
    End If
    
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub


Sub Articulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Articulo.Text <> "  -   -   " Then
        
            Articulo.Text = UCase(Articulo.Text)
            
            If Val(Solicitud.Text) <> 0 Then
                Sql1 = "Select *"
                Sql2 = " FROM Solic"
                Sql3 = " Where Solic.Solicitud = " + "'" + Solicitud.Text + "'"
                Sql4 = " and Solic.Articulo = " + "'" + Articulo.Text + "'"
                spSolic = Sql1 + Sql2 + Sql3 + Sql4
                Set rstSolic = db.OpenRecordset(spSolic, dbOpenSnapshot, dbSQLPassThrough)
                If rstSolic.RecordCount > 0 Then
                    rstSolic.Close
                        Else
                    T$ = "Ingreso de Aviso de No Entrega"
                    m$ = "No existe la M.P. en el numero de solicitud informado"
                    a% = MsgBox(m$, 0, T$)
                    Exit Sub
                End If
            End If
            
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + Articulo.Text + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesArticulo.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                FechaEntrega.SetFocus
                    Else
                m$ = "Materia Prima Inexistente"
                a% = MsgBox(m$, 0, "Ingreso de Aviso de Atraso de Entragado")
                Exit Sub
            End If
            
                Else
                
            FechaEntrega.SetFocus
            
        End If
    End If
    
    If KeyAscii = 27 Then
        Articulo.Text = "  -   -   "
        DesArticulo.Caption = ""
    End If
    
End Sub

Private Sub FechaEntrega_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(FechaEntrega.Text, Auxi)
        If Auxi = "S" Then
            Concepto.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        FechaEntrega.Text = "  /  /    "
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Concepto_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Fecha.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Productos"
    Opcion.AddItem "Materias Primas"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            Sql1 = "Select Codigo, Descripcion"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo >= " + "'" + "PT-00000-000" + "'"
            Sql4 = " and Terminado.Codigo <= " + "'" + "PT-999999-999" + "'"
            Sql5 = " Order by Codigo"
            spTerminado = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
            With rstTerminado
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstTerminado!Codigo + " " + rstTerminado!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstTerminado!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstTerminado.Close
            End If
            
            Sql1 = "Select Codigo, Descripcion"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo >= " + "'" + "DY-000-000" + "'"
            Sql4 = " and Articulo.Codigo <= " + "'" + "DW-9999-999" + "'"
            Sql5 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        WArticulo = Left$(rstArticulo!Codigo, 3) + "00" + Right$(rstArticulo!Codigo, 7)
                        IngresaItem = WArticulo + " " + rstArticulo!Descripcion
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
            
            
        Case 1
            Sql1 = "Select Codigo, Descripcion"
            Sql2 = " FROM Articulo"
            Sql3 = " Order by Codigo"
            spArticulo = Sql1 + Sql2 + Sql3
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
        
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub Pantalla_Click()
    pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Terminado.Text = WIndice.List(Indice)
            Call Terminado_Keypress(13)
            
        Case 1
            Indice = pantalla.ListIndex
            Articulo.Text = WIndice.List(Indice)
            Call Articulo_KeyPress(13)
        
        Case Else
    End Select
    
End Sub

Private Sub Form_Load()

    Concepto.Clear
    
    Concepto.AddItem ""
    Concepto.AddItem "Falta M.P.Local"
    Concepto.AddItem "Falta M.P. Importada"
    Concepto.AddItem "Cambio de Prioridades"
    Concepto.AddItem "Falta de Capacidad Disponible"
    Concepto.AddItem "Error del Sistema"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Problemas Vehiculos"
    Concepto.AddItem "Problemas Logistica"
    Concepto.AddItem "Problemas Recepcion Cliente"
    Concepto.AddItem "Varios"
    Concepto.AddItem "Corte de Luz"
    Concepto.AddItem "Pedido por el Cliente"
    Concepto.AddItem "Falta de Pago"
    Concepto.AddItem "Confirmacion Pedido Parcial"
    Concepto.AddItem "Envases"
    
    
    Concepto.ListIndex = 0

    Fecha.Text = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
    Pedido.Text = ""
    Cliente.Text = ""
    Terminado.Text = "  -     -   "
    Problema.Text = ""
    Articulo.Text = "  -   -   "
    FechaEntrega.Text = "  /  /    "
    DesCliente.Caption = ""
    DesTerminado.Caption = ""
    DesArticulo.Caption = ""
    
    If Val(WAtraso) <> 0 Then
    
        Sql1 = "Select *"
        Sql2 = " FROM Atraso"
        Sql3 = " Where Atraso.Numero = " + "'" + WAtraso + "'"
        spAtraso = Sql1 + Sql2 + Sql3
        Set rstAtraso = db.OpenRecordset(spAtraso, dbOpenSnapshot, dbSQLPassThrough)
        If rstAtraso.RecordCount > 0 Then
            Fecha.Text = rstAtraso!Fecha
            Pedido.Text = rstAtraso!Pedido
            Cliente.Text = rstAtraso!Cliente
            Terminado.Text = rstAtraso!Terminado
            Problema.Text = Trim(rstAtraso!Problema)
            Articulo.Text = rstAtraso!Articulo
            FechaEntrega.Text = rstAtraso!FechaEntrega
            Concepto.ListIndex = rstAtraso!Concepto
            rstAtraso.Close
        End If
        
        If Left$(Terminado.Text, 2) = "PT" Then
        
            Sql1 = "Select *"
            Sql2 = " FROM Terminado"
            Sql3 = " Where Terminado.Codigo = " + "'" + Terminado.Text + "'"
            spTerminado = Sql1 + Sql2 + Sql3
            Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
            If rstTerminado.RecordCount > 0 Then
                DesTerminado.Caption = rstTerminado!Descripcion
                rstTerminado.Close
            End If
            
                Else
                
            WArticulo = Left$(Terminado.Text, 3) + Right$(Terminado.Text, 7)
            Sql1 = "Select *"
            Sql2 = " FROM Articulo"
            Sql3 = " Where Articulo.Codigo = " + "'" + WArticulo + "'"
            spArticulo = Sql1 + Sql2 + Sql3
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                DesTerminado.Caption = rstArticulo!Descripcion
                rstArticulo.Close
            End If
            
        End If
            
        Sql1 = "Select *"
        Sql2 = " FROM Articulo"
        Sql3 = " Where Articulo.Codigo = " + "'" + Articulo.Text + "'"
        spArticulo = Sql1 + Sql2 + Sql3
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            DesArticulo.Caption = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        Sql1 = "Select *"
        Sql2 = " FROM Cliente"
        Sql3 = " Where Cliente.Cliente = " + "'" + Cliente.Text + "'"
        spCliente = Sql1 + Sql2 + Sql3
        Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        If rstCliente.RecordCount > 0 Then
            DesCliente.Caption = rstCliente!razon
            rstCliente.Close
        End If
        
        
    End If
        
End Sub

Private Sub Fecha_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Pedido_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub Terminado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Or KeyCode = 113 Or KeyCode = 114 Or KeyCode = 115 Or KeyCode = 116 Or KeyCode = 117 Or KeyCode = 118 Or KeyCode = 119 Or KeyCode = 120 Or KeyCode = 121 Then
        WFuncion = KeyCode
        Call Ejecuta_Funcion
    End If
End Sub

Private Sub Problema_KeyDown(KeyCode As Integer, Shift As Integer)
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

Private Sub FechaEntrega_KeyDown(KeyCode As Integer, Shift As Integer)
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





