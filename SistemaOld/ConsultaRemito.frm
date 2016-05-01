VERSION 5.00
Begin VB.Form PrgConsultaRemito 
   AutoRedraw      =   -1  'True
   Caption         =   "Consulta de Cuenta Corriente de Proveedores"
   ClientHeight    =   4650
   ClientLeft      =   420
   ClientTop       =   1410
   ClientWidth     =   11085
   LinkTopic       =   "Form2"
   ScaleHeight     =   4650
   ScaleWidth      =   11085
   Begin VB.ListBox PantaRemito 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   3240
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton ConsultaRemitoII 
      Caption         =   "Consulta Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8880
      TabIndex        =   11
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Height          =   540
      Left            =   8880
      TabIndex        =   10
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8880
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox Remito 
      Height          =   375
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   1
      Text            =   " "
      Top             =   3840
      Width           =   3375
   End
   Begin VB.TextBox Ayuda 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   8295
   End
   Begin VB.TextBox Proveedor 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1320
      MaxLength       =   11
      TabIndex        =   0
      Text            =   " "
      Top             =   3360
      Width           =   1695
   End
   Begin VB.ListBox Pantalla 
      Height          =   2595
      ItemData        =   "ConsultaRemito.frx":0000
      Left            =   120
      List            =   "ConsultaRemito.frx":0007
      TabIndex        =   3
      Top             =   600
      Width           =   8295
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
      Height          =   540
      Left            =   8880
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remito"
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
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Label DesProveedor 
      BackColor       =   &H00FFFFC0&
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
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3360
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proveedor"
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
      TabIndex        =   5
      Top             =   3360
      Width           =   1095
   End
End
Attribute VB_Name = "PrgConsultaRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RstProveedor As Recordset
Dim spProveedor As String
Dim cParam As String
Dim XParam As String

Dim ZZVector(1000) As String

Private Sub cmdClose_Click()
    PrgConsultaRemito.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Consulta_Click()

    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear
    
    spProveedor = "ListaProveedoresOrdConsulta"
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
        With RstProveedor
            .MoveFirst
            Do
                If .EOF = False Then
                    Auxi = !Proveedor
                    Call Ceros(Auxi, 11)
                    IngresaItem = Auxi + "      " + !Nombre
                    Pantalla.AddItem IngresaItem
                    IngresaItem = !Proveedor
                    WIndice.AddItem IngresaItem
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        RstProveedor.Close
    End If
            
    Rem Pantalla.Visible = True


End Sub

Private Sub ConsultaRemitoII_Click()
    
    Erase ZZVector
    WLugar = 0
    
    XEmpresa = WEmpresa
    If Val(XEmpresa) = 1 Then

        For Va = 1 To 7
    
            Select Case Va
                Case 1
                    WEmpresa = "0001"
                    txtOdbc = "Empresa01"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    WEmpresa = "0005"
                    txtOdbc = "Empresa05"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    WEmpresa = "0006"
                    txtOdbc = "Empresa06"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 5
                    WEmpresa = "0007"
                    txtOdbc = "Empresa07"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 6
                    WEmpresa = "0010"
                    txtOdbc = "Empresa10"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 7
                    WEmpresa = "0011"
                    txtOdbc = "Empresa11"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Renglon = " + "'" + "1" + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            ZSql = ZSql + " and Informe.FechaOrd >= " + "'" + "20100101" + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                With rstInforme
                    .MoveFirst
                    Do
                        If .EOF = False Then
                        
                            Zfecha1 = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            ZFecha2 = rstInforme!Fecha
                            ZZDife = DateDiff("d", ZFecha2, Zfecha1)
                            If ZZDife <= 90 Then
                                WLugar = WLugar + 1
                                ZZVector(WLugar) = rstInforme!Remito
                            End If
                            
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstInforme.Close
            End If
    
        Next Va
        
            Else
        
        For Va = 1 To 4
    
            Select Case Va
                Case 1
                    WEmpresa = "0002"
                    txtOdbc = "Empresa02"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 2
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 3
                    WEmpresa = "0008"
                    txtOdbc = "Empresa08"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case 4
                    WEmpresa = "0009"
                    txtOdbc = "Empresa09"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
            End Select
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Informe"
            ZSql = ZSql + " Where Informe.Renglon = " + "'" + "1" + "'"
            ZSql = ZSql + " and Informe.Proveedor = " + "'" + Proveedor.Text + "'"
            ZSql = ZSql + " and Informe.FechaOrd >= " + "'" + "20100101" + "'"
            spInforme = ZSql
            Set rstInforme = db.OpenRecordset(spInforme, dbOpenSnapshot, dbSQLPassThrough)
            If rstInforme.RecordCount > 0 Then
                With rstInforme
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            Zfecha1 = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                            ZFecha2 = rstInforme!Fecha
                            ZZDife = DateDiff("d", Now, LaFecha)
                            If ZZDife < 90 Then
                                WLugar = WLugar + 1
                                ZZVector(WLugar) = rstInforme!Remito
                            End If
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstInforme.Close
            End If
        
        Next Va
        
    End If
    
    Call Conecta_Empresa
    
    PantaRemito.Clear
    
    For Ciclo = 1 To WLugar
    
        ZZRemito = ZZVector(Ciclo)
        
                ZSql = ""
                ZSql = ZSql + "Select *"
                ZSql = ZSql + " FROM IvaComp"
                ZSql = ZSql + " Where IvaComp.Remito = " + "'" + ZZRemito + "'"
                ZSql = ZSql + " and IvaComp.Proveedor = " + "'" + Proveedor.Text + "'"
                spIvaComp = ZSql
                Set rstIvaComp = db.OpenRecordset(spIvaComp, dbOpenSnapshot, dbSQLPassThrough)
                If rstIvaComp.RecordCount > 0 Then
                    rstIvaComp.Close
                        Else
                    PantaRemito.AddItem ZZRemito
                End If
    
    Next Ciclo
    
    PantaRemito.Visible = True

End Sub

Private Sub Form_Activate()
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    Call Consulta_Click
End Sub

Private Sub Pantalla_Click()
    Indice = Pantalla.ListIndex
    Proveedor.Text = WIndice.List(Indice)
    Call Proveedor_KeyPress(13)
End Sub

Private Sub Form_Load()
    Proveedor.Text = ""
    DesProveedor.Caption = ""
    Remito.Text = ""
    Call Consulta_Click
End Sub

Private Sub PantaRemito_Click()
    ZZRemito = PantaRemito.Text
    If Trim(ZZRemito) <> "" Then
        If Trim(Remito.Text) = "" Then
            Remito.Text = ZZRemito
                Else
            Remito.Text = Remito.Text + "," + ZZRemito
        End If
    End If
    PantaRemito.Visible = False
    Call Remito_KeyPress(13)
End Sub

Private Sub Proveedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WProveedor = Proveedor.Text
        Proveedor.Text = WProveedor
        spProveedor = "ConsultaProveedores "
        XParam = "'" & Proveedor.Text & "'"
        Set RstProveedor = db.OpenRecordset(spProveedor + XParam, dbOpenSnapshot, dbSQLPassThrough)
        If RstProveedor.RecordCount > 0 Then
            DesProveedor.Caption = RstProveedor!Nombre
            RstProveedor.Close
            Remito.SetFocus
        End If
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Remito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Proveedor.SetFocus
        ZZPasaRemito = Remito.Text
        ZZPasaProveedor = Proveedor.Text
        ZZPasaProceso = 1
        PrgConsultaInforme.Show
    End If
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then

    Pantalla.Clear
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    If Ayuda.Text <> "" Then
        spProveedor = "ListaProveedoresOrdConsultaII " + "'" + Ayuda.Text + "'"
            Else
        spProveedor = "ListaProveedoresOrdConsulta"
    End If
    Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    If RstProveedor.RecordCount > 0 Then
    
    With RstProveedor
        .MoveFirst
        Do
            If .EOF = False Then
            
                da = Len(!Nombre) - WEspacios
                
                For aa = 1 To da
                    If Left$(UCase(Ayuda.Text), WEspacios) = Mid$(UCase(!Nombre), aa, WEspacios) Then
                    
                    
                        Auxi = Str$(!Proveedor)
                        Call Ceros(Auxi, 11)
                        IngresaItem = Auxi + "    " + !Nombre
                        Pantalla.AddItem IngresaItem
                        IngresaItem = !Proveedor
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
    
    RstProveedor.Close
    
    End If
    
    End If

End Sub

Private Sub Conecta_Empresa()

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
        Case Else
    End Select

End Sub

