VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form GRAFICO 
   Caption         =   "Form1"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox sector 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox planta 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   1035
      ItemData        =   "GRAFICO.frx":0000
      Left            =   1560
      List            =   "GRAFICO.frx":0016
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   1680
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox proye 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   5520
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "\\193.168.0.2\g$\vb\graficopta.rpt"
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
Attribute VB_Name = "GRAFICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spGraficinver As String
Rem Dim ver As Long
Dim wtotalreg As String

Dim ver55 As String
Dim d As Integer
Dim ver1 As Integer
Dim vector1(2, 20) As String

Dim vector(20, 20) As String


Private Sub Command1_Click()

        sql2 = " delete graficinver"
         spGraficinver = sql2
         Set rstgraficinver = db.OpenRecordset(spGraficinver, dbOpenSnapshot, dbSQLPassThrough)
 
 If sector.Text = "" Then
            m$ = "Se debe seleccionar el sector "
           A% = MsgBox(m$, 0, "Sector y planta")
              
          Else
          
          
          
          
          









   
 For i = 1 To wtotalreg
 
 vector(0, i) = 0
 
 Next
 
 
 
 ver45 = planta.Text
 ver56 = sector.Text
 

   
 Rem BUCLE DE 5 EMPRESAS
 
 
 
Rem For h = 1 To wtotalreg
  
 
 For j = 1 To 7

 
 planta.Text = j
 Gasto = 0
 

   ZSql = ""
    ZSql = ZSql + "Select sector, avance.importe"
    ZSql = ZSql + " FROM avance, proyecto"
    ZSql = ZSql + " Where avance.proyecto =  proyecto.codigo "
    ZSql = ZSql + " and proyecto.sector = " + "'" + sector.Text + "'"
    ZSql = ZSql + " and proyecto.planta = " + "'" + planta.Text + "'"
    ZSql = ZSql + " Order by avance.proyecto"
     spAvance = ZSql
    Set rstAvance = db.OpenRecordset(spAvance, dbOpenSnapshot, dbSQLPassThrough)
    If rstAvance.RecordCount > 0 Then
        With rstAvance
            .MoveFirst
            Do
                If .EOF = False Then
                
                       
                      sector2 = Int(Trim(rstAvance!sector))
                     Gasto = rstAvance!Importe
                      Rem buscar posicion
                      Call buscarposi(sector2, c)
                    Rem     d = Str$(c) + 1
                    
                    
                    
                    
                    
                    
                    
                    
                    vector(0, c) = vector(0, c) + Gasto
                    
               
               .MoveNext
               
               
               Else
                      Exit Do
               End If
            Loop
      End With
   rstAvance.Close
   
 End If




Select Case j
 
     Case 1
      
        pta1 = IIf(vector(0, c) = "", "0", vector(0, c))
        pta1 = Int(pta1)
    Case 2
        pta2 = IIf(vector(0, c) = "", "0", vector(0, c))
        pta2 = Int(pta2)
     Case 3
        pta3 = IIf(vector(0, c) = "", "0", vector(0, c))
        pta3 = Int(pta3)
    Case 4
         pta5 = IIf(vector(0, c) = "", "0", vector(0, c))
         pta5 = Int(pta5)
    Case 5
         pta6 = IIf(vector(0, c) = "", "0", vector(0, c))
         pta6 = Int(pta6)
    Case 6
         pta6 = IIf(vector(0, c) = "", "0", vector(0, c))
         pta6 = Int(pta7)
 
 
 End Select
 
 vector(0, c) = 0
 
 
 Next
 tot = Int(pta1 + pta2 + pta3 + pta5 + pta6 + pta7)
 

 
 
 
 
 
 
 ver = vector(c, c)
 ver12 = vector(1, 1)
 ver2 = vector(1, 1)





           

 
 For x = 1 To 6
  
Select Case x
 
 Case 1
  ver3 = pta1
  ver = Left(ver, 8) + " " + "Pta1"
 Case 2
   ver3 = pta2
 ver = Left(ver, 8) + " " + "Pta2"
 
 Case 3
    ver3 = pta3
 ver = Left(ver, 8) + " " + "Pta3"
 
 Case 4
    ver3 = pta5
 ver = Left(ver, 8) + " " + "Pta5"
Case 5
    ver3 = pta6
 ver = Left(ver, 8) + " " + "Pta6"
 Case 6
    ver3 = pta7
 ver = Left(ver, 8) + " " + "Pta7"
 
 
 End Select
 
 
 
Rem ver = IIf(vector(i, i) = "", "0", vector(i, i))
Rem ver3 = IIf(vector(0, i) = "", "0", vector(0, i))
Rem ver3 = Int(ver3)
         
           sql1 = "INSERT INTO graficinver ("
           sql2 = "sector,"
           sql3 = "Cantidad )"
           sql4 = " Values ("
           sql5 = "'" + ver + "',"
           sql6 = "'" + Str(ver3) + "')"
         
          spGraficinver = sql1 + sql2 + sql3 + sql4 + sql5 + sql6
           Set rstgraficinver = db.OpenRecordset(spGraficinver, dbOpenSnapshot, dbSQLPassThrough)
 Rem Next

 Rem          sql1 = "INSERT INTO graficinver ("
  Rem         sql2 = "sector,"
  Rem         sql3 = "Cantidad, "
  Rem         sql4 = "pta1,"
  Rem         sql5 = "pta2,"
  Rem         sql6 = "pta3,"
  Rem         sql7 = "pta5 ) "
  Rem         sql8 = " Values ("
  Rem         sql9 = "'" + ver + "',"
  Rem         sql10 = "'" + Str(tot) + "',"
  Rem         sql11 = "'" + Str(pta1) + "',"
  Rem         sql12 = "'" + Str(pta2) + "',"
  Rem         sql13 = "'" + Str(pta3) + "',"
  Rem         sql14 = "'" + Str(pta5) + "')"
  Rem         spGraficinver = sql1 + sql2 + sql3 + sql4 + sql5 + sql6 + sql7 + sql8 + sql9 + sql10 + sql11 + sql12 + sql13 + sql14
  Rem          Set rstgraficinver = db.OpenRecordset(spGraficinver, dbOpenSnapshot, dbSQLPassThrough)


 Next


Rem ver4 = 5
Rem sector2 = "INFRAESTRUCTURA"




Listado.Action = 1




End If

End Sub

Private Sub buscarposi(sector2, c)
B = 1
Do
If vector(B, 0) = sector2 Then
    c = B
    Exit Do
    Else
      B = B + 1
End If
Loop

End Sub

Private Sub Command2_Click()
 Listado.Action = 1
End Sub

Private Sub Form_Load()

            ZSql = ""
            ZSql = ZSql + "Select count(Codigo) as [totalreg]"
            ZSql = ZSql + " FROM sectorInve"
            spSectorInve = ZSql
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                rstSectorInve.MoveLast
                wtotalreg = IIf(IsNull(rstSectorInve!totalreg), "0", rstSectorInve!totalreg)
                rstSectorInve.Close
                   Else
              End If

           A = 1

            sql1 = "Select codigo,descripcion"
            sql2 = " FROM SectorInve"
            sql3 = " Order by SectorInve.Codigo"
            spSectorInve = sql1 + sql2 + sql3
            Set rstSectorInve = db.OpenRecordset(spSectorInve, dbOpenSnapshot, dbSQLPassThrough)
            If rstSectorInve.RecordCount > 0 Then
                With rstSectorInve
                    .MoveFirst
                    Do
                        If .EOF = False Then
   
                             If A <= wtotalreg Then
                                  vector(A, 0) = rstSectorInve!codigo
                                  vector(A, A) = rstSectorInve!descripcion
                                  List1.AddItem rstSectorInve!descripcion
                                  
                                 .MoveNext
                                  A = A + 1
                             End If
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstSectorInve.Close
   End If
   
 For i = 1 To wtotalreg
 
  vector(0, i) = 0
  Next
gastom = 0


End Sub

Private Sub List1_Click()
sector.Text = 1
sector.Text = sector.Text + List1.ListIndex
Rem ver22 = List1.List(List1.ListIndex)

Rem ver2 = List1.List(ver)


End Sub

Private Sub List2_Click()
planta.Text = 1
 planta.Text = planta.Text + List2.ListIndex

If List2.ListIndex = 3 Then
planta.Text = planta.Text + 1
End If
End Sub

