VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form miraproyecto 
   Caption         =   "Form1"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Ingreso 
      Caption         =   "Ingreso"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ListBox windice 
      Height          =   450
      Left            =   4320
      TabIndex        =   15
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid pantalla 
      Height          =   3135
      Left            =   1680
      TabIndex        =   14
      Top             =   2400
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5530
      _Version        =   327680
      Rows            =   100
      Cols            =   1
      FixedCols       =   0
      MergeCells      =   1
      AllowUserResizing=   2
   End
   Begin VB.TextBox Ayuda 
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
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   8175
   End
   Begin VB.TextBox descripcion 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox codigo 
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox proveedor 
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox provee 
      Height          =   495
      Left            =   6240
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox total 
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   9120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ComboBox WCombo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   6
      Top             =   3840
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
      Left            =   2040
      TabIndex        =   4
      Top             =   4320
      Width           =   375
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
      Left            =   840
      TabIndex        =   3
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton proceso 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox proyecto 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid WVector1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   19995
      _ExtentX        =   35269
      _ExtentY        =   8070
      _Version        =   327680
      BackColor       =   16777152
      Enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox WTexto3 
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Top             =   4320
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
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   1800
      MouseIcon       =   "miraproyecto.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "miraproyecto.frx":030A
      ToolTipText     =   "Consulta de Datos"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Gasto total"
      Height          =   255
      Left            =   7560
      TabIndex        =   11
      Top             =   6480
      Width           =   1215
   End
End
Attribute VB_Name = "miraproyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WFormato(20) As String
Dim WParametros(10, 20) As Double


Private Sub Consulta_Click()

    pantalla.Visible = False
    pantalla.Width = 8000
    pantalla.FixedCols = 0
    pantalla.Cols = 2
    pantalla.FixedRows = 1
    pantalla.Rows = 1000
    
    
    
    Rem WTitulo(1).Visible = False
    Rem WTitulo(2).Visible = False
    Rem Ayuda.Visible = False
   Rem  opcion.Clear

   Rem  opcion.AddItem "Proyectos"
   Rem  opcion.AddItem "Proveedores"

   Rem  opcion.Visible = True
     
   Rem On Error GoTo WError
    
   Rem opcion.Visible = False
     
    Dim IngresaItem As String

   Rem Call Limpia_Ayuda
    Lugarayuda = 0
    windice.Clear

  Rem  XIndice = opcion.ListIndex
    
    
    
   Rem Select Case XIndice
    Rem    Case 0
            
           Lugarayuda = 0
            sql1 = "Select *"
            Sql2 = " FROM Proyecto"
            Sql3 = " Order by Proyecto.Codigo"
            spProyecto = sql1 + Sql2 + Sql3
            Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
            If rstProyecto.RecordCount > 0 Then
                With rstProyecto
                    .MoveFirst
                    Do
                        
                        If .EOF = False Then
                            
                            Lugarayuda = Lugarayuda + 1
                            pantalla.Row = Lugarayuda
                            pantalla.Col = 0
                            pantalla.Text = rstProyecto!codigo
                            pantalla.Col = 1
                            pantalla.Text = rstProyecto!descripcion
                            IngresaItem = rstProyecto!codigo
                            windice.AddItem IngresaItem
                            
                            .MoveNext
                                
                                Else
                            Exit Do
                        End If
                  Loop
                End With
                rstProyecto.Close
            End If
            
     miraproyecto.Height = 7275
    miraproyecto.Width = 10590
    
    pantalla.ColWidth(1) = 6000
    pantalla.Visible = True
    Rem Ayuda.Visible = True
    Rem Ayuda.Text = ""
    Rem Ayuda.SetFocus
    
    
    
    
    Exit Sub
    
WError:
    Resume Next

End Sub

Private Sub pantalla1_Click()

    pantalla.Visible = False
    Ayuda.Visible = False
  Rem  WTitulo(1).Visible = False
   Rem WTitulo(2).Visible = False
    
Rem    Select Case XIndice
Rem        Case 0
Rem            Indice = Pantalla.Row - 1
Rem            Proyecto.Text = WIndice.List(Indice)
Rem            Call proyecto_KeyPress(13)
            
Rem        Case 1
Rem            Indice = Pantalla.Row - 1
  Rem          proveedor.Text = WIndice.List(Indice)
 Rem           Call Proveedor_Keypress(13)
            
miraproyecto.Height = 3000
miraproyecto.Width = 6000
pantalla.Visible = False
Rem        Case Else
Rem    End Select
    
End Sub

Private Sub Form_Load()
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
    WVector1.Visible = False
    total.Visible = False
    miraproyecto.Height = 3000
    miraproyecto.Width = 6000
pantalla.Visible = False

End Sub

Private Sub Ingreso_Click()
PrgAvance.Show
End Sub

Private Sub pantalla_Click()
Rem proyecto.Text = windice.List(Indice)
Indice = pantalla.Row - 1
            proyecto.Text = windice.List(Indice)
pantalla.Visible = False
Call proceso_click

End Sub

Private Sub proceso_click()
    
    descripcion.Text = ""

    total.Visible = True
    WVector1.Visible = True
    
    Gasto = 0
    total.Text = 0
    Call Limpia_Vector
    WRenglon = 0
    
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM proyecto"
    ZSql = ZSql + " Where proyecto.codigo = " + "'" + proyecto.Text + "'"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    If rstProyecto.RecordCount > 0 Then
        descripcion.Text = Trim(rstProyecto!descripcion)
         
        
        End If
    
    
    
    
    
    
    
    
    
    
    Rem fin nana
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM avance"
    ZSql = ZSql + " Where avance.proyecto = " + "'" + proyecto.Text + "'"
   Rem ZSql = ZSql " and Cronograma.Ano = " + "'" + Ano.Text + "'"
   ZSql = ZSql + " Order by avance.Codigo"
    
    spAvance = ZSql
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        With rstAvance
            .MoveFirst
            Do
                If .EOF = False Then
                
                    WRenglon = WRenglon + 1
                    WVector1.Row = WRenglon
                    Renglon = WRenglon
                     
                     codigo.Text = Trim(rstAvance!Tipo)
                               
   
   
   
   
   
   
                    WVector1.Col = 1
                    WVector1.Text = Trim(rstAvance!codigo)
                               
                
                    
                               
                    WVector1.Col = 2
                    WVector1.Text = Trim(rstAvance!descripcion)
            
                    
                    
                    provee.Text = Trim(rstAvance!proveedor)
            Rem busco el proveedor
                    
                    
         EmpresaReal = WEmpresa
         WEmpresa = "0001"
         txtOdbc = "Empresa01"
         strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
         Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        WEmpresa = EmpresaReal
         txtOdbc = "Empresa" + Right$(EmpresaReal, 2)
        
            Claveven$ = provee.Text
            spProveedor = "ConsultaProveedores " + "'" + Claveven$ + "'"
            Set rstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstProveedor.RecordCount > 0 Then
                    proveedor.Text = rstProveedor!nombre
                    rstProveedor.Close
              End If
              
Rem   strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
Rem            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                    
                    
                    WVector1.Col = 3
                    WVector1.Text = proveedor.Text
                    
Rem busco tipo
        sql1 = "Select *"
        Sql2 = " FROM SectorInve"
        Sql3 = " Where SectorInve.Codigo = " + "'" + codigo.Text + "'"
        spSectorInve = sql1 + Sql2 + Sql3
        Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
        If rstSectorInve.RecordCount > 0 Then
                   WVector1.Col = 4
                    WVector1.Text = Trim(rstAvance!descripcion)
        End If
                   
                   
                   
                   
                   Rem WVector1.Text = Trim(rstAvance!Proveedor)
                  Rem  WVector1.Text = Pusing("###,###.##", WVector1.Text)
                    
                  Rem  WVector1.Col = 4
                  Rem   WVector1.Text = Trim(rstAvance!Tipo)
                   Rem WVector1.Text = Pusing("###,###.##", WVector1.Text)
            
                  WVector1.Col = 5
                   WVector1.Text = Trim(rstAvance!Importe)
            
                     Gasto = Gasto + Trim(rstAvance!Importe)
                    
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstAvance.Close
    total.Text = Gasto
    
    End If
 
    
End Sub
Private Sub Limpia_Vector()

    WVector1.Clear

    Rem ponga la wvector1 en negritas
    WVector1.Font.Bold = True

    ' Inicalizo los Valores de las Variables
    
    
    
    
    miraproyecto.Height = 7275
    miraproyecto.Width = 10590
    
    
    
    
    
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
    WVector1.Rows = 101
    
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
                WVector1.Text = "codigo"
                WVector1.ColWidth(Ciclo) = 800
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 4
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 1
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 2
                WVector1.Text = "Descripcion"
                WVector1.ColWidth(Ciclo) = 2800
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 90
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 3
                WVector1.Text = "Proveedor"
                WVector1.ColWidth(Ciclo) = 2700
                WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
                WParametros(1, Ciclo) = 50
                WParametros(2, Ciclo) = 0
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
                WFormato(Ciclo) = ""
            Case 4
                WVector1.Text = "Tipo"
                WVector1.ColWidth(Ciclo) = 2500
                WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
                WParametros(1, Ciclo) = 15
                WParametros(2, Ciclo) = 1
                WParametros(3, Ciclo) = 0
                WParametros(4, Ciclo) = 0
               Rem WFormato(Ciclo) = "###.##"
       Case 5
               WVector1.Text = "Gasto"
        Rem        WVector1.ColWidth(Ciclo) = 1000
        Rem        WVector1.ColAlignment(Ciclo) = flexAlignRightCenter
        Rem        WParametros(1, Ciclo) = 4
        Rem        WParametros(2, Ciclo) = 1
        Rem        WParametros(3, Ciclo) = 1
        Rem        WParametros(4, Ciclo) = 0
        Rem        WFormato(Ciclo) = "###.##"
        Rem    Case 6
        Rem        WVector1.Text = "Observaciones"
        Rem        WVector1.ColWidth(Ciclo) = 2700
        Rem        WVector1.ColAlignment(Ciclo) = flexAlignLeftCenter
        Rem        WParametros(1, Ciclo) = 50
        Rem        WParametros(2, Ciclo) = 1
        Rem        WParametros(3, Ciclo) = 0
        Rem        WParametros(4, Ciclo) = 0
        Rem        WFormato(Ciclo) = ""
        End Select
    Next Ciclo
    
    Rem DESPILEGA LOS TITULOS
    
  Rem  WVector1.Row = 0
  Rem  For Ciclo = 1 To WVector1.Cols - 1
   Rem     WVector1.Col = Ciclo
        Rem WTitulo(Ciclo).Text = WVector1.Text
        Rem WTitulo(Ciclo).Left = WVector1.CellLeft + WVector1.Left
        Rem WTitulo(Ciclo).Top = WVector1.CellTop + WVector1.Top
        Rem WTitulo(Ciclo).Width = WVector1.CellWidth
        Rem WTitulo(Ciclo).Height = WVector1.CellHeight
        Rem WTitulo(Ciclo).Visible = True
Rem    Next Ciclo
    
    Rem CALCULA EL ANCHO TOTAL DE LA wvector1
    
    WAncho = 1200
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


Private Sub proyecto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
     proceso_click
  End If

End Sub

Private Sub WVector1_Click()
Rem proyecto.Text = windice.List(Indice)
proyecto.Text = pantalla.Row


End Sub
