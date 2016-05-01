VERSION 5.00
Begin VB.Form PrgLimpiaDatos 
   AutoRedraw      =   -1  'True
   Caption         =   "Reproceso de Stock de Materias Primas"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
End
Attribute VB_Name = "PrgLimpiaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Lugar1 As Integer
Private Lugar2 As Integer
Private Clave As String
Private WClave As String
Private WArticulo As String
Private WInicial As Double
Private WEntradas As Double
Private WSalidas As Double
Private WSaldo As Double
Private Vector(10000) As String
Dim Empe(12, 10) As String
Dim spArticulo As String
Dim rstArticulo As Recordset
Dim rstEstadistica As Recordset
Dim spEstadistica As String
Dim rstPedido As Recordset
Dim spPedido As String
Dim rstLaudo As Recordset
Dim rstOrden As Recordset
Dim spOrden As String
Dim spLaudo As String
Dim XParam As String
Private xLote(100, 7) As String

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
    Empe(9, 1) = "0009"
    Empe(9, 2) = "Empresa09"
    Empe(10, 1) = "0010"
    Empe(10, 2) = "Empresa10"
    Empe(11, 1) = "0011"
    Empe(11, 2) = "Empresa11"
    
    For A = 1 To 2
        
        WEmpresa = Empe(A, 1)
        txtOdbc = Empe(A, 2)
        strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
        Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        
        WFecha = "14/12/2001"
        WOrdFecha = Right$(WFecha, 4) + Mid$(WFecha, 4, 2) + Left$(WFecha, 2)
        XParam = "'" + WFecha + "','" _
                 + WOrdFecha + "'"
                 
        spArticulo = "ModificaArticuloFecha " + XParam
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    
        spTerminado = "ModificaTerminadoFecha " + XParam
        Set rstTerminado = db.OpenRecordset(spTerminado, dbOpenSnapshot, dbSQLPassThrough)

        Rem spPedido = "Limpia0"
        Rem Set rstPedido = db.OpenRecordset(spPedido, dbOpenSnapshot, dbSQLPassThrough)

        Rem spEstadistica = "Limpia1"
        Rem Set rstEstadistica = db.OpenRecordset(spEstadistica, dbOpenSnapshot, dbSQLPassThrough)

        Rem spArticulo = "Limpia2"
        Rem Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)

        Rem spOrden = "Limpia3"
        Rem Set rstOrden = db.OpenRecordset(spOrden, dbOpenSnapshot, dbSQLPassThrough)

        Rem spLaudo = "Limpia4"
        Rem Set rstLaudo = db.OpenRecordset(spLaudo, dbOpenSnapshot, dbSQLPassThrough)
        
    Next A
        
    
    Close
    End
    

End Sub

