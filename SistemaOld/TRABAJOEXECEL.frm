VERSION 5.00
Begin VB.Form TRABAJOEXCEL 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   3375
   End
End
Attribute VB_Name = "TRABAJOEXCEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objLibro As Object
Dim Ruta As String

Private Sub Command1_Click()

    Dim vbLine As Integer, X As Integer, vbOrden As Integer
    Dim vbSigue As Boolean
    Dim vbRow As Integer
    
    Set vbAplExc = CreateObject("Excel.application")

    Ruta = "c:\hoja.xls"

    vbAplExc.workbooks.Open ("C:\hoja.xls")
    Rem vbAplExc.Application.Visible = True
    
    Rem Min = 0
    Rem Max = vbAplExc.Range("B4").Value
    Rem ReDim vbOrdenes(vbAplExc.Range("B4").Value)
    Rem vbRow = vbAplExc.Range("B4").Value
    
    vbAplExc.Cells(4, 10).Value = "17.08.04"
    vbAplExc.Cells(6, 4).Value = "ALGO QUE YOP UIEROP"
    
    vbAplExc.Visible = True
    
    Rem With vbAplExc
    Rem     For vbLine = 1 To 6
    Rem         Select Case vbLine
    Rem             Case 4
    Rem                 fecha = .Cells(vbLine, 10).Value
    Rem
    Rem             Case 6
    Rem                 descri = .Cells(vbLine, 4).Value
    Rem             Case Else
    Rem         End Select
    Rem     Next vbLine
    Rem End With
                    
    Rem Stop
    
    Rem X = 1
    Rem With vbAplExc
    Rem     For vbLine = 7 To vbRow
    Rem         vbOrdenes(X).vbCodigo = .Cells(vbLine, 1).Value
    Rem         vbOrdenes(X).vbCantidad = .Cells(vbLine, 2).Value
    Rem         vbOrdenes(X).vbFecha = Format(.Cells(vbLine, 3).Value, "dd-mm-yyyy")
    Rem         vbOrdenes(X).VbUmedida = Trim(.Cells(vbLine, 4).Value)
    Rem         vbOrdenes(X).vbCentro = Trim(.Cells(vbLine, 5).Value)
    Rem         X = X + 1
    Rem     Next
    Rem     vbFecha = Format(.Cells(3, 2).Value, "dd-mm-yyyy")
    Rem End With
    Rem Consul = "select ord_num from orden_trabajo order by ord_num desc"
    Rem Set Rst1 = EterBase.OpenRecordset(Consul, dbOpenDynaset)
    Rem If (Rst1.EOF And Rst1.BOF) Then
    Rem     vbOrden = 1
    Rem         Else
    Rem     vbOrden = Rst1!ord_num
    Rem End If
    Rem Rst1.Close
    Rem prg.Visible = True
    Rem For X = 1 To vbAplExc.Range("B4").Val
    
        Rem vbAplExc.PrintOut Background
        Rem 'Comprobamos que Word no sigue imprimiendo
        Rem Do While vbAplExc.BackgroundPrintingStatus = 1
        Rem Loop
        
        Rem , vbAplExc.Close
        Rem vbAplExc.Printout
        Stop
        Rem vbAplExc.Application.DisplayAlerts = False
        Rem vbAplExc.Quit
        Rem Set vbAplExc = Nothing
    
    Close
    End
        


End Sub

