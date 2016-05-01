VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form prgestado 
   Caption         =   "listado de situacion de proyectos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Texto 
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox solicita 
      Height          =   315
      ItemData        =   "prgestado.frx":0000
      Left            =   1560
      List            =   "prgestado.frx":002B
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "todos"
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox estado 
      Height          =   315
      ItemData        =   "prgestado.frx":00AC
      Left            =   1560
      List            =   "prgestado.frx":00BF
      TabIndex        =   0
      Text            =   "todos"
      Top             =   360
      Width           =   2415
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   4680
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\\193.168.0.2\g$\system\proyectoestado.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      GroupSelectionFormula=   " "
      BoundReportFooter=   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
   End
   Begin VB.Label Label2 
      Caption         =   "SOLICITANTE"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "ESTADO"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "prgestado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ZVector(5000, 2) As String

  
Private Sub Command1_Click()




Select Case estado.ListIndex
Case 0
        Texto.Text = "1"
Case 1
        Texto.Text = "2"
Case 2
        Texto.Text = "3"
Case 3
         Texto.Text = "4"



End Select
  solicit = "'" + solicita.Text + "'"
  
  
   If estado.Text = "todos" Then
           If solicita.Text = "todos" Then
            
                      Rem pongo asterisco total
                 txtOdbc = "Empresa01"
                 strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                 Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                 DbConnect = db.Connect
                 DSQ = getDatabase(DbConnect)
               
                 Listado.SQLQuery = "SELECT Proyecto.Codigo,  Proyecto.Sector, Proyecto.Descripcion, Proyecto.presupuesto,Proyecto.Estado,Proyecto.FechaInicio,Proyecto.FechaFinal,Proyecto.Planta, Proyecto.Solicitante,Proyecto.gasto  " _
                 + "From " _
                 + DSQ + ".dbo.Proyecto Proyecto "
                
                 Listado.WindowTitle = "Proyectos en todos los estados"
                 Listado.Connect = Connect()
                 Listado.Action = 1
          
                    
             Else
                      Rem asterisco en estado
                                              
                        txtOdbc = "Empresa01"
                        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                        DbConnect = db.Connect
                        DSQ = getDatabase(DbConnect)
                         
                       Listado.SQLQuery = "SELECT Proyecto.Codigo,  Proyecto.Sector, Proyecto.Descripcion, Proyecto.presupuesto,Proyecto.Estado,Proyecto.FechaInicio,Proyecto.FechaFinal,Proyecto.Planta, Proyecto.Solicitante " _
                       + "From " _
                       + DSQ + ".dbo.Proyecto Proyecto " _
                       + "Where " _
                        + "Proyecto.solicitante = " + solicit
                      Rem Listado.ReportFileName = "proyectoestado.rpt"
                       Listado.Connect = Connect()
                       Listado.Action = 1

                Listado.WindowTitle = "Proyectos en todos los estados"
         End If
  
    Else
               If solicita.Text = "todos" Then
                    
                     Rem asterisco en solicitante
                     txtOdbc = "Empresa01"
                     strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                     Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                     DbConnect = db.Connect
                     
                     DSQ = getDatabase(DbConnect)
                       Listado.SQLQuery = "SELECT Proyecto.Codigo,  Proyecto.Sector, Proyecto.Descripcion, Proyecto.presupuesto,Proyecto.Estado,Proyecto.FechaInicio,Proyecto.FechaFinal,Proyecto.Planta, Proyecto.Solicitante " _
                       + "From " _
                       + DSQ + ".dbo.Proyecto Proyecto " _
                       + "Where " _
                       + "Proyecto.estado = " + Texto.Text
                       Rem Listado.ReportFileName = "proyectoestado.rpt"
                       Listado.Connect = Connect()
                       Listado.Action = 1
                       Listado.WindowTitle = "Proyectos " + estado.Text
                       

                          Else
                         

                              txtOdbc = "Empresa01"
                              strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                             Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    
                             DbConnect = db.Connect
                             DSQ = getDatabase(DbConnect)
                              Listado.SQLQuery = "SELECT Proyecto.Codigo,  Proyecto.Sector, Proyecto.Descripcion, Proyecto.presupuesto,Proyecto.Estado,Proyecto.FechaInicio,Proyecto.FechaFinal,Proyecto.Planta, Proyecto.Solicitante " _
                              + "From " _
                              + DSQ + ".dbo.Proyecto Proyecto " _
                              + "Where " _
                              + "Proyecto.estado = " + Texto.Text _
                              + "AND " _
                              + "Proyecto.solicitante = " + solicit
                              
                       Listado.WindowTitle = "Proyectos en " + estado.Text + " Para" + solicit
                              Listado.Connect = Connect()
                              Listado.Action = 1

             End If
End If











  







End Sub

Private Sub Form_Load()
 
 ZSql = ""
    ZSql = ZSql + "UPDATE Proyecto SET "
    ZSql = ZSql + " Proyecto.Gasto = 0 " + ","
    ZSql = ZSql + " Proyecto.Importe = 0 " + ","
    ZSql = ZSql + " Proyecto.Marca = " + "'" + "" + "'"
    spProyecto = ZSql
    Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Erase ZVector
    ZLugar = 0
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM Asigna"
 Rem   ZSql = ZSql + " Where Asigna.Ano = " + "'" + Ano.Text + "'"
    ZSql = ZSql + " Order by Asigna.Proyecto"
    spAsigna = ZSql
    Set rstAsigna = db.OpenRecordset(spAsigna, dbOpenSnapshot, dbSQLPassThrough)
    If rstAsigna.RecordCount > 0 Then
        With rstAsigna
            .MoveFirst
            Do
                If .EOF = False Then
                    ZLugar = ZLugar + 1
                    ZVector(ZLugar, 1) = rstAsigna!proyecto
                    ZVector(ZLugar, 2) = Str$(rstAsigna!Importe)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstAsigna.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZProyecto = ZVector(Ciclo, 1)
        ZGasto = 0
        ZImporte = Val(ZVector(Ciclo, 2))
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Avance"
        ZSql = ZSql + " Where Avance.Proyecto = " + "'" + ZProyecto + "'"
        spAvance = ZSql
        Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
        If rstAvance.RecordCount > 0 Then
            With rstAvance
                .MoveFirst
                Do
                    If .EOF = False Then
                        ZPeriodo = 0
                        If rstAvance!ordfecha >= "20070701" And rstAvance!ordfecha <= "20080631" Then
                            ZPeriodo = 1
                        End If
                        If rstAvance!ordfecha >= "20080701" And rstAvance!ordfecha <= "20090631" Then
                            ZPeriodo = 2
                        End If
                        If rstAvance!ordfecha >= "20090701" And rstAvance!ordfecha <= "20100631" Then
                            ZPeriodo = 3
                        End If
                        If rstAvance!ordfecha >= "20100701" And rstAvance!ordfecha <= "20110631" Then
                            ZPeriodo = 4
                        End If
                        If rstAvance!ordfecha >= "20110701" And rstAvance!ordfecha <= "20120631" Then
                            ZPeriodo = 5
                        End If
                        If rstAvance!ordfecha >= "20120701" And rstAvance!ordfecha <= "20130631" Then
                            ZPeriodo = 6
                        End If
                        If rstAvance!ordfecha >= "20130701" And rstAvance!ordfecha <= "20140631" Then
                            ZPeriodo = 7
                        End If
                        If rstAvance!ordfecha >= "20140701" And rstAvance!ordfecha <= "20150631" Then
                            ZPeriodo = 8
                        End If
                        If rstAvance!ordfecha >= "20150701" And rstAvance!ordfecha <= "20160631" Then
                            ZPeriodo = 9
                        End If
                        If rstAvance!ordfecha >= "20160701" And rstAvance!ordfecha <= "20070631" Then
                            ZPeriodo = 10
                        End If
                     Rem   If ZPeriodo = Val(Ano.Text) Then
                            ZGasto = ZGasto + rstAvance!Importe
                    Rem    End If
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstAvance.Close
        End If
        
        ZSql = ""
        ZSql = ZSql + "UPDATE Proyecto SET "
        ZSql = ZSql + " Proyecto.Gasto = " + "'" + Str$(ZGasto) + "',"
        ZSql = ZSql + " Proyecto.Importe = " + "'" + Str$(ZImporte) + "',"
        ZSql = ZSql + " Proyecto.Marca = " + "'" + "X" + "'"
        ZSql = ZSql + " Where Proyecto.Codigo = " + "'" + ZProyecto + "'"
        spProyecto = ZSql
        Set rstProyecto = db.OpenRecordset(spProyecto, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo

End Sub
