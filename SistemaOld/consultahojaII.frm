VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgConsultaHojaII 
   AutoRedraw      =   -1  'True
   Caption         =   "Asignacion de Hojas de Produccion"
   ClientHeight    =   7485
   ClientLeft      =   90
   ClientTop       =   690
   ClientWidth     =   11850
   LinkTopic       =   "Form2"
   ScaleHeight     =   7485
   ScaleWidth      =   11850
   Begin MSFlexGridLib.MSFlexGrid Muestra 
      Height          =   6255
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   11033
      _Version        =   327680
      Rows            =   4000
      Cols            =   5
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   10080
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "Ctacte.rpt"
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   5520
      MouseIcon       =   "consultahojaII.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "consultahojaII.frx":030A
      ToolTipText     =   "Salida"
      Top             =   6840
      Width           =   480
   End
End
Attribute VB_Name = "PrgConsultaHojaII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Dim XParam As String
Dim WGraba As String
Dim ZVector(100, 8) As String
Dim XEmpresa As String

Private Sub cmdClose_Click()
    PrgConsultaHojaII.Hide
    Unload Me
    Menu.Show
End Sub

Private Sub Form_Activate()
    Call Proceso_Click
End Sub

Private Sub Form_Load()

    Call Limpia_Vector
    
    Muestra.Font.Bold = True
    
    Muestra.ColWidth(0) = 200
    Muestra.ColWidth(1) = 1500
    Muestra.ColWidth(2) = 2000
    Muestra.ColWidth(3) = 2000
    Muestra.ColWidth(4) = 2000
    
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Poducto"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
End Sub

Private Sub Proceso_Click()

    XEmpresa = WEmpresa
    
    WSalida = "N"
        
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Producto"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
    Renglon = 0
    
    ZSql = ""
    ZSql = ZSql + "Select Hoja.Hoja, Hoja.Fecha, Hoja.Producto, Hoja.Teorico, Hoja.EstadoHoja, Hoja.Renglon"
    ZSql = ZSql + " FROM Hoja"
    ZSql = ZSql + " Where Hoja.EstadoHoja = 0 and Hoja.Renglon = 1"
    ZSql = ZSql + " Order by Hoja.Hoja"
            
    spHoja = ZSql
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
    
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                
                    Renglon = Renglon + 1
            
                    Muestra.TextMatrix(Renglon, 1) = Pusing("######", Str$(rstHoja!Hoja))
                    Muestra.TextMatrix(Renglon, 2) = rstHoja!Fecha
                    Muestra.TextMatrix(Renglon, 3) = rstHoja!Producto
                    Muestra.TextMatrix(Renglon, 4) = rstHoja!Teorico
                        
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    Muestra.Col = 0
    Muestra.Text = ""
    
    Muestra.TopRow = 1
    
End Sub

Private Sub Limpia_Vector()
    Muestra.Clear
    Muestra.Row = 0
    
    Muestra.Col = 1
    Muestra.Text = "Hoja"
    
    Muestra.Col = 2
    Muestra.Text = "Fecha"
    
    Muestra.Col = 3
    Muestra.Text = "Producto"
    
    Muestra.Col = 4
    Muestra.Text = "Cantidad"
    
End Sub

Private Sub Muestra_DblClick()

    Muestra.Col = 1
    ZHojaProceso = Muestra.Text
    Muestra.Col = 3
    ZTerminadoProceso = Muestra.Text
    Muestra.Col = 4
    ZCantidadProceso = Muestra.Text
    
    Rem PrgConsultaHoja.Hide
    Rem Unload Me
    PrgHojaNuevaII.Show
    
End Sub

