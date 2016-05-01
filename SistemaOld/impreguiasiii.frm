VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgImpreGuiaSiii 
   AutoRedraw      =   -1  'True
   Caption         =   "Proceso de Impresion de Pedidos"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
   Begin VB.Frame Aviso 
      BackColor       =   &H00FF8080&
      Height          =   4455
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton Cancela 
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   2040
         TabIndex        =   3
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "EXISTEN SOLICITUD DE GUIAS DE TRASLADO INTERNO PENDIENTE DE REALIZAR"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   480
         TabIndex        =   2
         Top             =   360
         Width           =   5055
      End
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Aceptar"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   1200
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "WPedPen.rpt"
      Destination     =   1
      WindowTitle     =   "Listado de Iva ventas"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
End
Attribute VB_Name = "PrgImpreGuiaSiii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstSolguia As Recordset
Dim spSolguia As String

Private Sub Acepta_Click()

    XEmpresa = "5"

    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
    WLugar = 0
        
    XParam = "'" + "N" + " '"
    spSolguia = "ListaSolGuiaPendiente " + XParam
    Set rstSolguia = db.OpenRecordset(spSolguia, dbOpenSnapshot, dbSQLPassThrough)
    If rstSolguia.RecordCount > 0 Then
            
        With rstSolguia
    
            .MoveFirst
            If .NoMatch = False Then
                Do
                
                    Pasa = "N"
                    Select Case Val(XEmpresa)
                        Case 1
                            If Val(rstSolguia!Desde) = 1 Then
                                Pasa = "S"
                            End If
                        Case 3
                            If Val(rstSolguia!Desde) = 2 Then
                                Pasa = "S"
                            End If
                        Case 5
                            If Val(rstSolguia!Desde) = 3 Then
                                Pasa = "S"
                            End If
                        Case 6
                            If Val(rstSolguia!Desde) = 4 Then
                                Pasa = "S"
                            End If
                        Case 7
                            If Val(rstSolguia!Desde) = 5 Then
                                Pasa = "S"
                            End If
                        Case Else
                    End Select
                            
                    If Pasa = "S" Then
                        WLugar = WLugar + 1
                    End If
                    
                    .MoveNext
                
                    If .EOF = True Then
                        Exit Do
                    End If
                
                Loop
            End If
        
        End With
        
        rstSolguia.Close
        
    End If
    
    If WLugar > 0 Then
        PrgImpreGuiaSiii.Refresh
        Aviso.Visible = True
        Aviso.Refresh
        For A = 1 To 10
            Beep
        Next A
        PrgImpreGuiaSiii.Refresh
        Aviso.Visible = True
        Aviso.Refresh
            Else
        Call Cancela_click
    End If

End Sub

Private Sub Cancela_click()
    PrgImpreGuiaSiii.Hide
    Unload Me
    End
End Sub

Private Sub Form_Load()
    Call Acepta_Click
End Sub

Private Sub Form_Activate()
    OPEN_FILE_Empresa
End Sub

