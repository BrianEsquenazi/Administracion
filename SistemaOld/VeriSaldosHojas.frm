VERSION 5.00
Begin VB.Form PrgVeriSaldosHojas 
   AutoRedraw      =   -1  'True
   Caption         =   "Control de Ordenes de Compra Pendientes"
   ClientHeight    =   5205
   ClientLeft      =   1890
   ClientTop       =   1080
   ClientWidth     =   8145
   LinkTopic       =   "Form2"
   ScaleHeight     =   5205
   ScaleWidth      =   8145
End
Attribute VB_Name = "PrgVeriSaldosHojas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WArticulo As String
Private WClave As String
Dim rstHoja As Recordset
Dim spHoja As String
Dim rstTerminado As Recordset
Dim spTerminado As String
Dim XParam As String
Dim Vector(10000, 4) As String
Dim Empe(10, 10) As String
Private WDescripcion As String
Private WSaldo As Double
Private XSaldo As Double

Private Sub Form_Load()
    Call Proceso_Click
End Sub

Private Sub Proceso_Click()

    Empe(1, 1) = "0001"
    Empe(1, 2) = "Empresa01"
    Empe(2, 1) = "0002"
    Empe(2, 2) = "Empresa02"
    Empe(3, 1) = "0003"
    Empe(3, 2) = "Empresa03"
    Empe(4, 1) = "0004"
    Empe(4, 2) = "Empresa04"
    Empe(5, 1) = "0005"
    Empe(5, 2) = "Empresa05"
    Empe(6, 1) = "0006"
    Empe(6, 2) = "Empresa06"
    Empe(7, 1) = "0007"
    Empe(7, 2) = "Empresa07"
    Empe(8, 1) = "0008"
    Empe(8, 2) = "Empresa08"
    
    For A = 1 To 8
        
    WEmpresa = Empe(A, 1)
    txtOdbc = Empe(A, 2)
    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)

    Erase Vector
    Renglon = 0
    
    spTerminado = "ModificaTerminadoProceso0"
    Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
    
    spHoja = "ListaHojaReproceso"
    Set rstHoja = db.OpenRecordset(spHoja, dbOpenSnapshot, dbSQLPassThrough)
    If rstHoja.RecordCount > 0 Then
        With rstHoja
            .MoveFirst
            Do
                If .EOF = False Then
                    Renglon = Renglon + 1
                    aa = rstHoja!hoja
                    Vector(Renglon, 1) = rstHoja!Producto
                    Vector(Renglon, 2) = Str$(rstHoja!teorico)
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstHoja.Close
    End If
    
    For Ciclo = 1 To Renglon
    
        WProducto = Vector(Ciclo, 1)
        WTeorico = Vector(Ciclo, 2)
    
        XParam = "'" + WProducto + "','" _
                 + WTeorico + "'"
        spTerminado = "ModificaTerminadoProceso " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    Next A
    
    PrgVeriSaldosHojas.Hide
    Unload Me
    Close
    End
    
End Sub

