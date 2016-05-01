VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form PrgGrabaIb 
   AutoRedraw      =   -1  'True
   Caption         =   "Carga de % de Ingresos Brutos"
   ClientHeight    =   4605
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   4605
   ScaleWidth      =   8145
   Begin VB.TextBox LIsta 
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Cancela 
      Caption         =   "Cancela"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Acepta 
      Caption         =   "Acepta"
      Height          =   615
      Left            =   3000
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   7320
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "PedpenII.rpt"
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
Attribute VB_Name = "PrgGrabaIb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Uno As String
Private Dos As String
Private Tres As String
Dim rstcliente As Recordset
Dim spCliente As String
Dim rstLiberaTerminado As Recordset
Dim spLiberaTerminado As String
Dim XParam As String
Dim WVector(1000, 5) As String
Dim ZCarga(10) As String
Dim LugarVector As Integer
Dim WTipopro As String
Dim WDesdeFec As String
Dim WHastaFec As String
Dim LugarLibera As Integer
Dim LugarLiberaI As Integer
Dim LugarLiberaII As Integer
Dim LugarLiberaIII As Integer
Dim ZCampo(20) As String

Dim WPorceI As String
Dim WPorceII As String


Dim WWPorceI As String
Dim WWPorceII As String

Private Sub Acepta_Click()


Rem *****************************************************
Rem *PARAMETROS DE EJECUCION                            *
Rem *****************************************************
Rem ***Padron Embargo***                                *
Rem ***  Si  a=1 No a=0                                 *
     a = 1
     
Rem ***Padron provincia Bs As***                        *
Rem ***  Si  p=1 No p=0                                 *
     p = 1
     
Rem ***Padron caba magnitudes superadas Y ALTO RIESGO***              *
Rem ***Padron alto riesgo***                            *
Rem ***  Si ms=1  No ms=0                               *
     ms = 1



   
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    


    Rem
    Rem Padron Embargo
    Rem
     
     If a = 1 Then
     
         Open "c:\padron\Padron_Embargo.txt" For Input As #10
    
         ZSql = ""
         ZSql = ZSql + "UPDATE Proveedor SET "
         ZSql = ZSql + "Embargo = " + "'" + "" + "'"
         spProveedor = ZSql
         Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    
      Do
           Line Input #10, WDatos
         If EOF(10) Then Exit Do
        
           WCuit = Mid$(WDatos, 9, 11)
           WCuitBusqueda = Mid$(WDatos, 11, 8)
           ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "Embargo = " + "'" + "S" + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        Loop Until EOF(10)
        
        Close #10
    
   End If
    
    
    
   

    Rem
    Rem Padron provincia Bs As
    Rem
    If p = 1 Then
    
         Open "c:\padron\padronrete.txt" For Input As #10
         Rem Open "c:\padron\DADA.txt" For Input As #10
    
         Rem ZSql = ""
         Rem ZSql = ZSql + "UPDATE Cliente SET "
         Rem ZSql = ZSql + "PorceIb = " + "'" + "6" + "'"
         Rem spCliente = ZSql
         Rem Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
         
         ZSql = ""
         ZSql = ZSql + "UPDATE Proveedor SET "
         ZSql = ZSql + "PorceIb = " + "'" + "4" + "'"
         spProveedor = ZSql
         Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
         ZSuma = 0
         aa = Time
      Do
    
         Line Input #10, WDatos
          If EOF(10) Then Exit Do
        
           WCuit = Mid$(WDatos, 30, 11)
        
           WCuitBusqueda = Mid$(WDatos, 32, 8)
           WPorceI = Mid$(WDatos, 48, 1) + "." + Mid$(WDatos, 50, 2)
        
            ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
            ZSuma = ZSuma + 1
            ZSumaII = ZSumaII + 1
        
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Cliente SET "
            Rem ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            Rem ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            Rem spCliente = ZSql
            Rem Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
             If ZSumaII > 1000 Then
                LIsta.Text = ZSuma
                DoEvents
                ZSumaII = 0
             End If
        
      Loop Until EOF(10)
    
    Close #10
   
  End If
   
   
   
   
   
   
   
   
   
   
   
   
   
   
    Rem
    Rem Padron provincia Bs As
    Rem
    If p = 1 Then
    
         Open "c:\padron\padronperce.txt" For Input As #10
    
         ZSql = ""
         ZSql = ZSql + "UPDATE Cliente SET "
         ZSql = ZSql + "PorceIb = " + "'" + "8" + "'"
         spCliente = ZSql
         Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
         
         Rem ZSql = ""
         Rem ZSql = ZSql + "UPDATE Proveedor SET "
         Rem ZSql = ZSql + "PorceIb = " + "'" + "3" + "'"
         Rem spProveedor = ZSql
         Rem Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
         ZSuma = 0
         aa = Time
      Do
    
         Line Input #10, WDatos
          If EOF(10) Then Exit Do
        
           WCuit = Mid$(WDatos, 30, 11)
        
           WCuitBusqueda = Mid$(WDatos, 32, 8)
           WPorceI = Mid$(WDatos, 48, 1) + "." + Mid$(WDatos, 50, 2)
        
            ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
            ZSuma = ZSuma + 1
            ZSumaII = ZSumaII + 1
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spCliente = ZSql
            Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Proveedor SET "
            Rem ZSql = ZSql + "PorceIb = " + "'" + WPorceII + "'"
            Rem ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            Rem spProveedor = ZSql
            Rem Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
             If ZSumaII > 1000 Then
                LIsta.Text = ZSuma
                DoEvents
                ZSumaII = 0
             End If
        
      Loop Until EOF(10)
    
    Close #10
   
  End If
   
   
   
    Rem
    Rem Padron caba magnitudes superadas
    Rem

  If ms = 1 Then
  
      If Val(WEmpresa) = 1 Then

            Open "c:\padron\padron_magnitudes.txt" For Input As #10
        
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "PorceIbCabaAnterior = PorceIbCaba"
            spCliente = ZSql
            Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIbCabaAnterior = PorceIbCaba"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "PorceIbCaba = " + "'" + "0" + "'"
            spCliente = ZSql
            Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIbCaba = " + "'" + "0" + "'"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
            ZSuma = 0
            ZSumaII = 0
            aa = Time

             Do
             
                 Line Input #10, WDatos
                 If EOF(10) Then Exit Do
                 
                 Erase ZCampo
                 NroCampo = 0
                 ZDesde = 1
                 ZZHasta = Len(WDatos)
                 For Ciclo = 1 To ZZHasta
                     If Mid$(WDatos, Ciclo, 1) = ";" Then
                         NroCampo = NroCampo + 1
                         ZLargo = Ciclo - ZDesde
                         ZCampo(NroCampo) = Mid$(WDatos, ZDesde, ZLargo)
                         ZDesde = Ciclo + 1
                     End If
                 Next Ciclo
                 
                 WCuit = ZCampo(1)
                 ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
                 
                 WPorceI = ZCampo(3)
                 WPorceII = ZCampo(3)
                 Call Convierte1_datos(WPorceI, WWPorceI)
                 Call Convierte1_datos(WPorceII, WWPorceII)
                 
                 ZSuma = ZSuma + 1
                 ZSumaII = ZSumaII + 1
                 
                 ZSql = ""
                 ZSql = ZSql + "UPDATE Cliente SET "
                 ZSql = ZSql + "PorceIbCaba = " + "'" + WWPorceI + "'"
                 ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
                 spCliente = ZSql
                 Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
                 
                 ZSql = ""
                 ZSql = ZSql + "UPDATE Proveedor SET "
                 ZSql = ZSql + "PorceIbCaba = " + "'" + WWPorceII + "'"
                 ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
                 spProveedor = ZSql
                 Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                 
                 If ZSumaII > 1000 Then
                     LIsta.Text = ZSuma
                     DoEvents
                     ZSumaII = 0
                 End If
             
            Loop Until EOF(10)
    
            Close #10
    
    
            Open "c:\padron\padron_altoriesgo.txt" For Input As #10
        
            ZSuma = 0
            ZSumaII = 0
            aa = Time
        
            Do
        
                Line Input #10, WDatos
                If EOF(10) Then Exit Do
            
                Erase ZCampo
                NroCampo = 0
                ZDesde = 1
                ZZHasta = Len(WDatos)
                For Ciclo = 1 To ZZHasta
                    If Mid$(WDatos, Ciclo, 1) = ";" Then
                        NroCampo = NroCampo + 1
                        ZLargo = Ciclo - ZDesde
                        ZCampo(NroCampo) = Mid$(WDatos, ZDesde, ZLargo)
                        ZDesde = Ciclo + 1
                    End If
                Next Ciclo
            
                WCuit = ZCampo(4)
                ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
                
                WPorceI = ZCampo(8)
                WPorceII = ZCampo(9)
                Call Convierte1_datos(WPorceI, WWPorceI)
                Call Convierte1_datos(WPorceII, WWPorceII)
                
                ZSuma = ZSuma + 1
                ZSumaII = ZSumaII + 1
                
                ZSql = ""
                ZSql = ZSql + "UPDATE Cliente SET "
                ZSql = ZSql + "PorceIbCaba = " + "'" + WWPorceI + "'"
                ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
                spCliente = ZSql
                Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            
                ZSql = ""
                ZSql = ZSql + "UPDATE Proveedor SET "
                ZSql = ZSql + "PorceIbCaba = " + "'" + WWPorceII + "'"
                ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
                spProveedor = ZSql
                Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
                
                If ZSumaII > 1000 Then
                    LIsta.Text = ZSuma
                    DoEvents
                    ZSumaII = 0
                End If
        
            Loop Until EOF(10)
        
            Close #10
    
        End If

    End If
    


   
    WEmpresa = "0001"
    txtOdbc = "Empresa01"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
    
    

   

    Rem
    Rem Padron provincia Bs As
    Rem
    If p = 1 Then
    
         Open "c:\padron\padronrete.txt" For Input As #10
         Rem Open "c:\padron\DADA.txt" For Input As #10
    
         Rem ZSql = ""
         Rem ZSql = ZSql + "UPDATE Cliente SET "
         Rem ZSql = ZSql + "PorceIb = " + "'" + "6" + "'"
         Rem spCliente = ZSql
         Rem Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
         
         ZSql = ""
         ZSql = ZSql + "UPDATE Proveedor SET "
         ZSql = ZSql + "PorceIb = " + "'" + "4" + "'"
         spProveedor = ZSql
         Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
         ZSuma = 0
         aa = Time
      Do
    
         Line Input #10, WDatos
          If EOF(10) Then Exit Do
        
           WCuit = Mid$(WDatos, 30, 11)
        
           WCuitBusqueda = Mid$(WDatos, 32, 8)
           WPorceI = Mid$(WDatos, 48, 1) + "." + Mid$(WDatos, 50, 2)
        
            ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
            ZSuma = ZSuma + 1
            ZSumaII = ZSumaII + 1
        
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Cliente SET "
            Rem ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            Rem ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            Rem spCliente = ZSql
            Rem Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
             If ZSumaII > 1000 Then
                LIsta.Text = ZSuma
                DoEvents
                ZSumaII = 0
             End If
        
      Loop Until EOF(10)
    
    Close #10
   
  End If
   
   
   
   
   
   
   
   
   
   
   
   
   
   
    Rem
    Rem Padron provincia Bs As
    Rem
    If p = 1 Then
    
         Open "c:\padron\padronperce.txt" For Input As #10
    
         ZSql = ""
         ZSql = ZSql + "UPDATE Cliente SET "
         ZSql = ZSql + "PorceIb = " + "'" + "8" + "'"
         spCliente = ZSql
         Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
         
         Rem ZSql = ""
         Rem ZSql = ZSql + "UPDATE Proveedor SET "
         Rem ZSql = ZSql + "PorceIb = " + "'" + "3" + "'"
         Rem spProveedor = ZSql
         Rem Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
         ZSuma = 0
         aa = Time
      Do
    
         Line Input #10, WDatos
          If EOF(10) Then Exit Do
        
           WCuit = Mid$(WDatos, 30, 11)
        
           WCuitBusqueda = Mid$(WDatos, 32, 8)
           WPorceI = Mid$(WDatos, 48, 1) + "." + Mid$(WDatos, 50, 2)
        
            ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
            ZSuma = ZSuma + 1
            ZSumaII = ZSumaII + 1
        
            ZSql = ""
            ZSql = ZSql + "UPDATE Cliente SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spCliente = ZSql
            Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
        
            Rem ZSql = ""
            Rem ZSql = ZSql + "UPDATE Proveedor SET "
            Rem ZSql = ZSql + "PorceIb = " + "'" + WPorceII + "'"
            Rem ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            Rem spProveedor = ZSql
            Rem Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        
             If ZSumaII > 1000 Then
                LIsta.Text = ZSuma
                DoEvents
                ZSumaII = 0
             End If
        
      Loop Until EOF(10)
    
    Close #10
   
  End If
   
       
    
    
    
    
    
    
    
    Call Cancela_click
    
  End Sub



Private Sub AceptaII_Click()



    Rem WEmpresa = "0001"
    Rem txtOdbc = "Empresa01"
    Rem strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Rem Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Rem Open "c:\padron\Padron_Embargo.txt" For Input As #10
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Proveedor SET "
    Rem ZSql = ZSql + "Embargo = " + "'" + "" + "'"
    Rem spProveedor = ZSql
    Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    
    Rem Do
    Rem     Line Input #10, WDatos
    Rem     If EOF(10) Then Exit Do
        
    Rem     wcuit = Mid$(WDatos, 9, 11)
    Rem     wcuitBusqueda = Mid$(WDatos, 11, 8)
        
    Rem     WProveedor = ""
        
    Rem     ZSql = ""
    Rem     ZSql = ZSql + "Select *"
    Rem     ZSql = ZSql + " FROM Proveedor"
    Rem     ZSql = ZSql + " Where Proveedor.Cuit LIKE " + "'" + "%" + wcuitBusqueda + "%" + "'"
    Rem     ZSql = ZSql + " Order by Cuit"
    Rem     spProveedor = ZSql
    Rem     Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If RstProveedor.RecordCount > 0 Then
    Rem         With RstProveedor
    Rem             .MoveFirst
    Rem             Do
    Rem                 If .EOF = False Then
    Rem                     WProveedor = RstProveedor!Proveedor
    Rem                      .MoveNext
    Rem                         Else
    Rem                     Exit Do
    Rem                 End If
    Rem             Loop
    Rem         End With
    Rem         RstProveedor.Close
    Rem     End If
        
    Rem     If WProveedor <> "" Then
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "UPDATE Proveedor SET "
    Rem         ZSql = ZSql + "Embargo = " + "'" + "S" + "'"
    Rem         ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
    Rem         spProveedor = ZSql
    Rem         Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem     End If
    Rem
    Rem Loop Until EOF(10)
    Rem
    Rem Close #10
    
    
    
    

    Rem Open "c:\padron\padron.txt" For Input As #10
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Cliente SET "
    Rem ZSql = ZSql + "PorceIb = " + "'" + "6" + "'"
    Rem spCliente = ZSql
    Rem Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Proveedor SET "
    Rem ZSql = ZSql + "PorceIb = " + "'" + "3" + "'"
    Rem spProveedor = ZSql
    Rem Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    Rem ZSuma = 0
    Rem aa = Time
    
    Rem Do
    Rem     Line Input #10, WDatos
    Rem     If EOF(10) Then Exit Do
    Rem
    Rem     wcuit = Mid$(WDatos, 28, 11)
    Rem     wcuitBusqueda = Mid$(WDatos, 30, 8)
    Rem     wporceI = Mid$(WDatos, 46, 1) + "." + Mid$(WDatos, 48, 2)
    Rem     wporceII = Mid$(WDatos, 51, 1) + "." + Mid$(WDatos, 53, 2)
    Rem
    Rem
    Rem
    Rem     WCliente = ""
    Rem     WClienteI = ""
    Rem     WProveedor = ""
    Rem
    Rem     ZSuma = ZSuma + 1
    Rem
    Rem     ZSql = ""
    Rem     ZSql = ZSql + "Select *"
    Rem     ZSql = ZSql + " FROM Cliente"
    Rem     ZSql = ZSql + " Where Cliente.Cuit LIKE " + "'" + "%" + wcuitBusqueda + "%" + "'"
    Rem     ZSql = ZSql + " Order by Cuit"
    Rem     spCliente = ZSql
    Rem     Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If rstCliente.RecordCount > 0 Then
    Rem         With rstCliente
    Rem             .MoveFirst
    Rem             Do
    Rem                 If .EOF = False Then
    Rem                     If WCliente = "" Then
    Rem                         WCliente = rstCliente!Cliente
    Rem                             Else
    Rem                         WClienteI = rstCliente!Cliente
    Rem                     End If
    Rem                      .MoveNext
    Rem                         Else
    Rem                     Exit Do
    Rem                 End If
    Rem             Loop
    Rem         End With
    Rem         rstCliente.Close
    Rem     End If
    Rem
    Rem     If WCliente <> "" Then
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "UPDATE Cliente SET "
    Rem         ZSql = ZSql + "PorceIb = " + "'" + wporceI + "'"
    Rem         ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
    Rem         spCliente = ZSql
    Rem         Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Rem     End If
    Rem
    Rem     If WClienteI <> "" Then
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "UPDATE Cliente SET "
    Rem         ZSql = ZSql + "PorceIb = " + "'" + wporceI + "'"
    Rem         ZSql = ZSql + " Where Cliente = " + "'" + WClienteI + "'"
    Rem         spCliente = ZSql
    Rem         Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    Rem     End If
    Rem
    Rem     ZSql = ""
    Rem     ZSql = ZSql + "Select *"
    Rem     ZSql = ZSql + " FROM Proveedor"
    Rem     ZSql = ZSql + " Where Proveedor.Cuit LIKE " + "'" + "%" + wcuitBusqueda + "%" + "'"
    Rem     ZSql = ZSql + " Order by Cuit"
    Rem     spProveedor = ZSql
    Rem     Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem     If RstProveedor.RecordCount > 0 Then
    Rem         With RstProveedor
    Rem             .MoveFirst
    Rem             Do
    Rem                 If .EOF = False Then
    Rem                     WProveedor = RstProveedor!Proveedor
    Rem                      .MoveNext
    Rem                         Else
    Rem                     Exit Do
    Rem                 End If
    Rem             Loop
    Rem         End With
    Rem         RstProveedor.Close
    Rem     End If
    Rem
    Rem     If WProveedor <> "" Then
    Rem         ZSql = ""
    Rem         ZSql = ZSql + "UPDATE Proveedor SET "
    Rem         ZSql = ZSql + "PorceIb = " + "'" + wporceII + "'"
    Rem         ZSql = ZSql + " Where Proveedor = " + "'" + WProveedor + "'"
    Rem         spProveedor = ZSql
    Rem         Set RstProveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    Rem     End If
    Rem
    Rem Loop Until EOF(10)
    Rem
    Rem Close #10
    
    
    
    
    
    
    
    
    
    

    WEmpresa = "0008"
    txtOdbc = "Empresa08"
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Close #10
    Open "c:\padron\padron.txt" For Input As #10
    
    Rem ZSql = ""
    Rem ZSql = ZSql + "UPDATE Cliente SET "
    Rem ZSql = ZSql + "PorceIb = " + "'" + "6" + "'"
    Rem spCliente = ZSql
    Rem Set rstCliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "UPDATE Proveedor SET "
    ZSql = ZSql + "PorceIb = " + "'" + "4" + "'"
    spProveedor = ZSql
    Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSuma = 0
    aa = Time
    
    Do
        Line Input #10, WDatos
        If EOF(10) Then Exit Do
        
        WCuit = Mid$(WDatos, 28, 11)
        
        WCuitBusqueda = Mid$(WDatos, 30, 8)
        WPorceI = Mid$(WDatos, 46, 1) + "." + Mid$(WDatos, 48, 2)
        WPorceII = Mid$(WDatos, 51, 1) + "." + Mid$(WDatos, 53, 2)
        
        ZZCuit = Left$(WCuit, 2) + "-" + Mid$(WCuit, 3, 8) + "-" + Mid$(WCuit, 11, 1)
        
        WCliente = ""
        WClienteI = ""
        WProveedor = ""
        
        ZSuma = ZSuma + 1
        ZSumaII = ZSumaII + 1
        
        If dada = 999 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Cliente"
            ZSql = ZSql + " Where Cliente.Cuit LIKE " + "'" + "%" + WCuitBusqueda + "%" + "'"
            ZSql = ZSql + " Order by Cuit"
            spCliente = ZSql
            Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            If rstcliente.RecordCount > 0 Then
                With rstcliente
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            If WCliente = "" Then
                                WCliente = rstcliente!Cliente
                                    Else
                                WClienteI = rstcliente!Cliente
                            End If
                             .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstcliente.Close
            End If
            
            If WCliente <> "" Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Cliente SET "
                ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
                ZSql = ZSql + " Where Cliente = " + "'" + WCliente + "'"
                spCliente = ZSql
                Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            End If
            
            If WClienteI <> "" Then
                ZSql = ""
                ZSql = ZSql + "UPDATE Cliente SET "
                ZSql = ZSql + "PorceIb = " + "'" + WPorceI + "'"
                ZSql = ZSql + " Where Cliente = " + "'" + WClienteI + "'"
                spCliente = ZSql
                Set rstcliente = db.OpenRecordset(spCliente, dbOpenSnapshot, dbSQLPassThrough)
            End If
        
        End If
        
        If dada = 999 Then
        
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Proveedor"
            ZSql = ZSql + " Where Proveedor.Cuit LIKE " + "'" + "%" + WCuitBusqueda + "%" + "'"
            ZSql = ZSql + " Order by Cuit"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
            If rstproveedor.RecordCount > 0 Then
                With rstproveedor
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            WProveedor = rstproveedor!Proveedor
                             .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstproveedor.Close
            End If
        End If
        
        Rem If WProveedor <> "" Then
            ZSql = ""
            ZSql = ZSql + "UPDATE Proveedor SET "
            ZSql = ZSql + "PorceIb = " + "'" + WPorceII + "'"
            ZSql = ZSql + " Where Cuit = " + "'" + ZZCuit + "'"
            spProveedor = ZSql
            Set rstproveedor = db.OpenRecordset(spProveedor, dbOpenSnapshot, dbSQLPassThrough)
        Rem End If
        
        If ZSumaII > 1000 Then
            LIsta.Text = ZSuma
            DoEvents
            ZSumaII = 0
        End If
        
        
    Loop Until EOF(10)
    
    Call Cancela_click
    
End Sub




Private Sub Cancela_click()
    PrgGrabaIb.Hide
    Unload Me
    Close
    End
End Sub


