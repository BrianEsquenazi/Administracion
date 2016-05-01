VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCambiaParametro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Valores Standard de Especificaciones de Pt"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   495
   ClientWidth     =   15270
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form2"
   ScaleHeight     =   8340
   ScaleWidth      =   15270
   Visible         =   0   'False
   Begin VB.ComboBox Tipo 
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
      Left            =   1320
      TabIndex        =   21
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox WTexto2 
      Alignment       =   1  'Right Justify
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
      Height          =   285
      Left            =   1800
      TabIndex        =   19
      Top             =   3480
      Width           =   375
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   1800
      TabIndex        =   18
      Top             =   4080
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.TextBox WTexto1 
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2400
      TabIndex        =   17
      Top             =   3480
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
      Index           =   6
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   2640
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
      Index           =   5
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2520
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
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   2400
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2280
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
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2400
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   8493
      _Version        =   327680
      BackColor       =   16777152
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
      Left            =   10200
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Ensayo 
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
      Left            =   1320
      MaxLength       =   6
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   9120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   6120
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Limpia 
      Caption         =   "Limpia Pantalla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   6720
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Graba 
      Caption         =   "Graba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7920
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   5880
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
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
      Height          =   2220
      ItemData        =   "CambiaParametro.frx":0000
      Left            =   3360
      List            =   "CambiaParametro.frx":0007
      TabIndex        =   1
      Top             =   5880
      Visible         =   0   'False
      Width           =   8055
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "impreped.rpt"
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Top             =   3480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      _Version        =   327680
      BackColor       =   16776960
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
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label DesEnsayo 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "PrgCambiaParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String

Dim rstEnsayos As Recordset
Dim spEnsayos As String

Dim XParam As String
Dim ZZVector(10000, 10) As String
Dim ZZVerifica(10000) As String
Dim Renglon As Integer

Rem para el vector

Dim WBorra(10000, 20) As String
Dim WParametros(10, 20) As Double
Dim WFormato(20) As String
Dim WControl As String

Private Sub cmdClose_Click()
    PrgCambiaParametro.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

    If Val(EmpresaActual) <> 0 Then
        XEmpresa = EmpresaActual
        Call Conecta_Empresa
    End If

End Sub

Private Sub Graba_Click()

    Rem On Error GoTo WError
    
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
    
    
    For A = 1 To Renglon
        
        WProducto = WVector1.TextMatrix(A, 1)
        WValor = Left$(WVector1.TextMatrix(A, 2), 50)
        WValor1 = Left$(WVector1.TextMatrix(A, 5), 50)
        
        ZZRenglon = ZZVector(A, 1)
    
        Select Case ZZRenglon
            Case 1
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor1 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor11 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 2
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor2 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor22 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 3
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor3 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor33 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 4
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor4 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor44 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 5
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor5 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor55 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 6
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor6 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor66 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 7
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor7 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor77 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 8
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor8 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor88 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 9
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor9 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor99 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            Case 10
                ZSql = ""
                ZSql = ZSql & "UPDATE EspecifUnifica SET "
                ZSql = ZSql & "Valor10 = " + "'" + WValor + "',"
                ZSql = ZSql & "Valor1010 = " + "'" + WValor1 + "'"
                ZSql = ZSql & " Where Producto = " + "'" + WProducto + "'"
                        
                spEspecifUnifica = ZSql
                Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            Case Else
        End Select
    Next A
    
    Call Conecta_Empresa
        
    Call Limpia_Click
        
    Exit Sub

WError:
     Resume Next
        
End Sub

Private Sub Limpia_Click()

    Ensayo.Text = ""
    DesEnsayo.Caption = ""
    
    Call Limpia_Vector
    
    Ensayo.SetFocus

End Sub

Private Sub Form_Load()

    Tipo.Clear
    
    Tipo.AddItem ""
    Tipo.AddItem "Quimicos"
    Tipo.AddItem "Pigmentos"
    Tipo.AddItem "Colorantes"
    Tipo.AddItem "Resto"
    
    Tipo.ListIndex = 0
    

    Ensayo.Text = ""
    DesEnsayo.Caption = ""
    
    Call Limpia_Vector
     
End Sub

Private Sub Proceso_Click()

    Call Limpia_Vector
    Erase ZZVector
    
    Renglon = 0
    
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
    
    
    
    
    ZZTipo = Tipo.ListIndex
    Erase ZZVerifica
    ZZLugar = 0
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Filtro"
    spFiltro = ZSql
    Set rstFiltro = db.OpenRecordset(spFiltro, dbOpenSnapshot, dbSQLPassThrough)
    If rstFiltro.RecordCount > 0 Then
    
        With rstFiltro
            .MoveFirst
            Do
                If .EOF = False Then
                
                    If ZZTipo = 0 Or ZZTipo = 4 Or ZZTipo = rstFiltro!Tipo Then
                        ZZLugar = ZZLugar + 1
                        ZZVerifica(ZZLugar) = rstFiltro!Codigo
                    End If
        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstFiltro.Close
    End If
    
    
    
    
    
    ZSql = ""
    ZSql = ZSql + "Select EspecifUnifica.Producto, EspecifUnifica.Ensayo1, EspecifUnifica.Valor1, EspecifUnifica.Desde1, EspecifUnifica.Hasta1, EspecifUnifica.Valor11, EspecifUnifica.Ensayo2, EspecifUnifica.Valor2, EspecifUnifica.Desde2, EspecifUnifica.Hasta2, EspecifUnifica.Valor22, EspecifUnifica.Ensayo3, EspecifUnifica.Valor3, EspecifUnifica.Desde3, EspecifUnifica.Hasta3, EspecifUnifica.Valor33, EspecifUnifica.Ensayo4, EspecifUnifica.Valor4, EspecifUnifica.Desde4, EspecifUnifica.Hasta4, EspecifUnifica.Valor44, EspecifUnifica.Ensayo5, EspecifUnifica.Valor5, EspecifUnifica.Desde5, EspecifUnifica.Hasta5, EspecifUnifica.Valor55, EspecifUnifica.Ensayo6, EspecifUnifica.Valor6, EspecifUnifica.Desde6, EspecifUnifica.Hasta6, EspecifUnifica.Valor66, EspecifUnifica.Ensayo7, EspecifUnifica.Valor7, EspecifUnifica.Desde7, EspecifUnifica.Hasta7, EspecifUnifica.Valor77, EspecifUnifica.Ensayo8, EspecifUnifica.Valor8, EspecifUnifica.Desde8, EspecifUnifica.Hasta8, EspecifUnifica.Valor88, "
    ZSql = ZSql + "EspecifUnifica.Ensayo9, EspecifUnifica.Valor9, EspecifUnifica.Desde9, EspecifUnifica.Hasta9, EspecifUnifica.Valor99, EspecifUnifica.Ensayo10, EspecifUnifica.Valor10, EspecifUnifica.Desde10, EspecifUnifica.Hasta10, EspecifUnifica.Valor1010 "
    ZSql = ZSql + " FROM EspecifUnifica"
    ZSql = ZSql + " Order by Producto"
    spEspecifUnifica = ZSql
    Set rstEspecifUnifica = db.OpenRecordset(spEspecifUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecifUnifica.RecordCount > 0 Then
    
        With rstEspecifUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Rem If UCase(rstEspecifUnifica!Producto) = "PT-09830-100" Then Stop
                                
                    If Left$(UCase(rstEspecifUnifica!Producto), 2) = "PT" Then
                        
                        If ZZTipo = 4 Then
                            ZZEntra = "S"
                            For CicloVeri = 1 To ZZLugar
                                If Trim(UCase(rstEspecifUnifica!Producto)) = Trim(UCase(ZZVerifica(CicloVeri))) Then
                                    ZZEntra = "N"
                                    Exit For
                                End If
                            Next CicloVeri
                                Else
                            ZZEntra = "N"
                            For CicloVeri = 1 To ZZLugar
                                If Trim(UCase(rstEspecifUnifica!Producto)) = Trim(UCase(ZZVerifica(CicloVeri))) Then
                                    ZZEntra = "S"
                                    Exit For
                                End If
                            Next CicloVeri
                        End If
                            
                        If ZZEntra = "S" Then
                    
                            If rstEspecifUnifica!Ensayo1 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo1
                                ZZValor = IIf(IsNull(rstEspecifUnifica!Valor1), "", rstEspecifUnifica!Valor1)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde1), "", rstEspecifUnifica!Desde1)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta1), "", rstEspecifUnifica!Hasta1)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor11), "", rstEspecifUnifica!Valor11)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 1
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo2 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo2
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor2), "", rstEspecifUnifica!valor2)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde2), "", rstEspecifUnifica!Desde2)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta2), "", rstEspecifUnifica!Hasta2)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor22), "", rstEspecifUnifica!Valor22)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 2
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo3 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo3
                                ZZValor = IIf(IsNull(rstEspecifUnifica!Valor3), "", rstEspecifUnifica!Valor3)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde3), "", rstEspecifUnifica!Desde3)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta3), "", rstEspecifUnifica!Hasta3)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor33), "", rstEspecifUnifica!Valor33)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 3
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo4 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo4
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor4), "", rstEspecifUnifica!valor4)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde4), "", rstEspecifUnifica!Desde4)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta4), "", rstEspecifUnifica!Hasta4)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor44), "", rstEspecifUnifica!Valor44)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 4
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo5 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo5
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor5), "", rstEspecifUnifica!valor5)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde5), "", rstEspecifUnifica!Desde5)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta5), "", rstEspecifUnifica!Hasta5)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor55), "", rstEspecifUnifica!Valor55)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 5
                                
                            End If
                            
                            
                            If rstEspecifUnifica!Ensayo6 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo6
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor6), "", rstEspecifUnifica!valor6)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde6), "", rstEspecifUnifica!Desde6)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta6), "", rstEspecifUnifica!Hasta6)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor66), "", rstEspecifUnifica!Valor66)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 6
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo7 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo7
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor7), "", rstEspecifUnifica!valor7)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde7), "", rstEspecifUnifica!Desde7)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta7), "", rstEspecifUnifica!Hasta7)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor77), "", rstEspecifUnifica!Valor77)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 7
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo8 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo8
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor8), "", rstEspecifUnifica!valor8)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde8), "", rstEspecifUnifica!Desde8)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta8), "", rstEspecifUnifica!Hasta8)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor88), "", rstEspecifUnifica!Valor88)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 8
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo9 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo9
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor9), "", rstEspecifUnifica!valor9)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde9), "", rstEspecifUnifica!Desde9)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta9), "", rstEspecifUnifica!Hasta9)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor99), "", rstEspecifUnifica!Valor99)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 9
                                
                            End If
                            
                            If rstEspecifUnifica!Ensayo10 = Val(Ensayo.Text) Then
                            
                                Renglon = Renglon + 1
                                
                                ZZEnsayo = rstEspecifUnifica!Ensayo10
                                ZZValor = IIf(IsNull(rstEspecifUnifica!valor10), "", rstEspecifUnifica!valor10)
                                ZZDesde = IIf(IsNull(rstEspecifUnifica!Desde10), "", rstEspecifUnifica!Desde10)
                                ZZHasta = IIf(IsNull(rstEspecifUnifica!Hasta10), "", rstEspecifUnifica!Hasta10)
                                ZZValorI = IIf(IsNull(rstEspecifUnifica!Valor1010), "", rstEspecifUnifica!Valor1010)
                                            
                                WVector1.TextMatrix(Renglon, 1) = rstEspecifUnifica!Producto
                                WVector1.TextMatrix(Renglon, 2) = ZZValor
                                WVector1.TextMatrix(Renglon, 3) = Trim(ZZDesde)
                                WVector1.TextMatrix(Renglon, 4) = Trim(ZZHasta)
                                WVector1.TextMatrix(Renglon, 5) = ZZValorI
                                
                                ZZVector(Renglon, 1) = 10
                                
                            End If
                            
                        End If
                                            
                    End If
        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecifUnifica.Close
    End If
    
    Call Conecta_Empresa
    
    WVector1.TopRow = 1
    WVector1.Row = 1
    WVector1.Col = 1
    
    Call StartEdit
    
    Graba.Enabled = True

End Sub

Private Sub Ensayo_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    
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
        
        spEnsayo = "Select * FROM Ensayos Where Codigo = " + Ensayo.Text
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            DesEnsayo.Caption = Trim(rstEnsayo!Descripcion)
            rstEnsayo.Close
        End If
        
        Call Conecta_Empresa
                
        Call Proceso_Click
                
        WVector1.TopRow = 1
        WVector1.Row = 1
        WVector1.Col = 2
        Call StartEdit
            
    End If
End Sub



Private Sub Consulta_Click()

    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    Rem XIndice = Opcion.ListIndex
    XIndice = 0
    
    Select Case XIndice
        Case 0
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
        
            spEnsayo = "ListaEnsayos"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                With rstEnsayo
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                            Enspantalla.AddItem IngresaItem
                            IngresaItem = rstEnsayo!Codigo
                            EnsIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstEnsayo.Close
            End If
            
            Call Conecta_Empresa
            
            
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    
    pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Ensayo.Text = WIndice.List(Indice)
            Call Ensayo_keypress(13)
            
        Case Else
    End Select
    
End Sub


Rem
Rem Controles de la wvector1
Rem

Private Sub GridEditText(ByVal KeyAscii As Integer)

    XColumna = WVector1.Col
    XTipoDato = WParametros(3, XColumna)

    Select Case XTipoDato
        Case 0
            WTexto1.Left = WVector1.CellLeft + WVector1.Left
            WTexto1.Top = WVector1.CellTop + WVector1.Top
            WTexto1.Width = WVector1.CellWidth
            WTexto1.Height = WVector1.CellHeight
            WTexto1.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto1.Text = WVector1.Text
                    WTexto1.SelStart = Len(WTexto1.Text)
                Case Else
                    WTexto1.Text = Chr$(KeyAscii)
                    WTexto1.SelStart = 1
            End Select
            WTexto1.Visible = True
            WTexto1.SetFocus
        Case 1
            WTexto2.Left = WVector1.CellLeft + WVector1.Left
            WTexto2.Top = WVector1.CellTop + WVector1.Top
            WTexto2.Width = WVector1.CellWidth
            WTexto2.Height = WVector1.CellHeight
            WTexto2.MaxLength = WParametros(1, XColumna)
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    WTexto2.Text = WVector1.Text
                    Rem WTexto2.SelStart = Len(WTexto2.Text)
                    WTexto2.SelStart = 0
                Case Else
                    WTexto2.Text = Chr$(KeyAscii)
                    WTexto2.SelStart = 1
            End Select
            WTexto2.Visible = True
            WTexto2.SetFocus
        Case 2
            WTexto3.Left = WVector1.CellLeft + WVector1.Left
            WTexto3.Top = WVector1.CellTop + WVector1.Top
            WTexto3.Width = WVector1.CellWidth
            WTexto3.Height = WVector1.CellHeight
            Select Case KeyAscii
                Case 0 To Asc(" ")
                    If Len(WVector1.Text) = 10 Then
                        WTexto3.Text = WVector1.Text
                            Else
                        WTexto3.Text = "  /  /    "
                    End If
                    WTexto3.SelStart = 0
                Case Else
                    WTexto3.Text = Chr$(KeyAscii) + " /  /    "
                    WTexto3.SelStart = 1
            End Select
            WTexto3.Visible = True
            WTexto3.SetFocus
        Case Else
    End Select

End Sub

Private Sub EndEdit()
    Pasa = 0
    If WCombo1.Visible Then
        Pasa = 0
        WVector1.Text = WCombo1.Text
        WCombo1.Visible = False
            Else
        If WTexto1.Visible Then
            Pasa = 1
            WVector1.Text = WTexto1.Text
            WTexto1.Visible = False
                Else
            If WTexto2.Visible Then
                Pasa = 1
                WVector1.Text = WTexto2.Text
                WTexto2.Visible = False
                    Else
                If WTexto3.Visible Then
                    Pasa = 1
                    WVector1.Text = WTexto3.Text
                    WTexto3.Visible = False
                End If
            End If
        End If
    End If
    If Pasa = 1 Then
        If WFormato(WVector1.Col) <> "" Then
            WVector1.Text = Pusing(WFormato(WVector1.Col), WVector1.Text)
        End If
        Rem Call Suma_Datos
    End If
End Sub

Private Sub GridEditCombo()
    ' Position the ComboBox over the cell.
    WCombo1.Left = WVector1.CellLeft + WVector1.Left
    WCombo1.Top = WVector1.CellTop + WVector1.Top
    WCombo1.Width = WVector1.CellWidth
    WCombo1.Visible = True
    WCombo1.SetFocus
End Sub

Private Sub WTexto1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto1.Text = ""
            
        Rem F1
        Case 113
            WTexto1.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 123
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Col > 1 Then
                WVector1.Col = WVector1.Col - 1
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            WTexto2.Text = ""
            
        Rem F1
        Case 113
            WTexto2.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            DoEvents
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
                Call StartEdit
            End If
    
        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                Rem End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                Rem End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

Private Sub WTexto3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            ' Leave the text unchanged.
            WTexto3.Text = "  /  /    "
            
        Rem F1
        Case 113
            WTexto3.Text = WVector1.Text

        Case vbKeyReturn
            ' Finish editing.
            WVector1.SetFocus
            Call Control_Campo
            If WControl = "S" Then
                Call Control_wvector1
            End If
            Call StartEdit

        Case vbKeyDown
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row < WVector1.Rows - 1 Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row + 1
                End If
            End If
            Call StartEdit

        Case vbKeyUp
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.Row > WVector1.FixedRows Then
                Call Control_Campo
                If WControl = "S" Then
                    WVector1.Row = WVector1.Row - 1
                End If
            End If
            Call StartEdit
        Case 34
            ' Move down 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow < WVector1.Rows - 12 Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow + 12
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit
            
        Case 33
            ' Move up 1 row.
            WVector1.SetFocus
            DoEvents
            If WVector1.TopRow - 12 > WVector1.FixedRows Then
                Rem Call Control_Campo
                Rem If WControl = "S" Then
                    WVector1.TopRow = WVector1.TopRow - 12
                    WVector1.Row = WVector1.TopRow
                        Else
                    WVector1.TopRow = 1
                    WVector1.Row = WVector1.TopRow
                Rem End If
            End If
            Call StartEdit

    End Select
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto2_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

' Do not beep on Return or Escape.
Private Sub WTexto3_KeyPress(KeyAscii As Integer)
    If (KeyAscii = vbKeyReturn) Or _
       (KeyAscii = vbKeyEscape) _
            Then KeyAscii = 0
End Sub

' Make the change.
Private Sub WCombo1_Click()
    WVector1.SetFocus
End Sub


Private Sub WVector1_Click()
    StartEdit
End Sub

Private Sub WVector1_LeaveCell()
    EndEdit
End Sub

Private Sub WVector1_GotFocus()
    EndEdit
End Sub

Private Sub WVector1_KeyPress(KeyAscii As Integer)
    XColumna = WVector1.Col
    Select Case WParametros(4, WVector1.Col)
        Case 1
        Case Else
            If WParametros(2, XColumna) = 0 Then
                GridEditText KeyAscii
            End If
    End Select
End Sub


Rem
Rem Desde aca empieza las rutinas a cambiar
Rem

Private Sub StartEdit()
    Select Case WParametros(4, WVector1.Col)
        Case 1
            Rem Carga los datos en el caso que el campo a editar sea un combo
            WCombo1.Clear
            WCombo1.AddItem "Campo1"
            WCombo1.AddItem "Campo2"
            On Error Resume Next
            WCombo1.Text = WVector1.Text
            On Error GoTo 0
            GridEditCombo
        Case Else
            If WParametros(2, WVector1.Col) = 0 Then
                GridEditText Asc(" ")
            End If
    End Select
End Sub

Private Sub Control_wvector1()
    Select Case WVector1.Col
        Case 5
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 5
        Case 2
            If WVector1.Row < WVector1.Rows - 1 Then
                WVector1.Row = WVector1.Row + 1
            End If
            WVector1.Col = 2
        Case Else
            If WVector1.Col < WVector1.Cols - 1 Then
                WVector1.Col = WVector1.Col + 1
            End If
    End Select
    WVector1.SetFocus
    GridEditText KeyAscii
End Sub

Private Sub Control_Campo()
    XColumna = WVector1.Col
    XFila = WVector1.Row
    WControl = "S"
    Select Case XColumna
            
        Case Else
            WVector1.Col = XColumna
    End Select
End Sub

Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    WTexto1.FontName = WVector1.FontName
    WTexto1.FontSize = WVector1.FontSize
    WTexto1.Visible = False
    WTexto2.FontName = WVector1.FontName
    WTexto2.FontSize = WVector1.FontSize
    WTexto2.Visible = False
    WTexto3.FontName = WVector1.FontName
    WTexto3.FontSize = WVector1.FontSize
    WTexto3.Visible = False
    WCombo1.FontName = WVector1.FontName
    WCombo1.FontSize = WVector1.FontSize
    WCombo1.Visible = False

    ' Establesco loa Valores de la wvector1
    
    WVector1.FixedCols = 1
    WVector1.Cols = 6
    WVector1.FixedRows = 1
    WVector1.Rows = 10001
    
    Rem Descripcion de los datos a Informar
    
    Rem Titulo
    Rem WVector1.Text = "Articulo"
    
    Rem Longitud
    Rem WVector1.ColWidth(Ciclo) = 1200
    
    Rem Alineacion de la columna
    Rem WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
    
    Rem cantidad maxima de caracteres
    Rem WParametros(1, 1) = 4

    Rem indica si el campo es editable
    Rem (0 es editable, 1 no es editable)
    Rem WParametros(2, 1) = 0
    
    Rem tipo de datos del ingreso
    Rem (0 si es texto, 1 si es numerico, 2 si es fecha)
    Rem WParametros(3, 1) = 0
    
    Rem SI ES TEXTO O COMBO
    Rem (0 si es texto, 1 SI ES COMBO)
    Rem WParametros(4, 1) = 0
    
    Rem Descripcion de los datos a Informar
    
    WVector1.ColWidth(0) = 400
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Select Case Ciclo
            Case 1
                WVector1.Text = "Producto"
                WVector1.ColWidth(Ciclo) = 1400
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 12
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Desde"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Hasta"
                WVector1.ColWidth(Ciclo) = 1000
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 10
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 5
                WVector1.Text = "Descripcion II"
                WVector1.ColWidth(Ciclo) = 5000
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
    WVector1.Row = 0
    For Ciclo = 1 To WVector1.Cols - 1
        WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 340
    For Ciclo = 0 To WVector1.Cols - 1
        WAncho = WAncho + WVector1.ColWidth(Ciclo)
    Next Ciclo
    Rem WVector1.Width = WAncho

    ' Size the columns.
    Font.Name = WVector1.Font.Name
    Font.Size = WVector1.Font.Size
    
    Rem Parametro que indica que el usuario puede
    Rem modificar el tamaño de las celdas
    WVector1.AllowUserResizing = flexResizeBoth
    
    WVector1.Col = 1
    WVector1.Row = 1
    
End Sub

Private Sub WVector1_Scroll()
    WTexto1.Visible = False
    WTexto2.Visible = False
    WTexto3.Visible = False
End Sub
































