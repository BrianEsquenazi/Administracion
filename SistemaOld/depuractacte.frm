VERSION 5.00
Begin VB.Form PrgDepuraCtaCte 
   AutoRedraw      =   -1  'True
   Caption         =   "Depuracion de centavos en la cuenta corriente de Clientes y Proveedores"
   ClientHeight    =   6405
   ClientLeft      =   1410
   ClientTop       =   1155
   ClientWidth     =   9585
   LinkTopic       =   "Form2"
   ScaleHeight     =   6405
   ScaleWidth      =   9585
   Begin VB.CommandButton Cancelar 
      Caption         =   "Cancelar"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
   End
   Begin VB.CommandButton Aceptar 
      Caption         =   "Aceptar"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   1440
      Width           =   3135
   End
End
Attribute VB_Name = "PrgDepuraCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCtaCte As Recordset
Dim spCtaCte As String
Dim rstCtaCtePrv As Recordset
Dim spCtaCtePrv As String

Dim XParam As String

Private Sub Cancelar_Click()

    PrgDepuraCtaCte.Hide
    Unload Me
    Menu.Show
    
End Sub

Private Sub Aceptar_Click()

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCte SET "
     ZSql = ZSql + " Saldo = 0"
    ZSql = ZSql + " Where Saldo > -0.1 and Saldo < 0.1"
    spCtaCte = ZSql
    Set rstCtaCte = db.OpenRecordset(spCtaCte, dbOpenSnapshot, dbSQLPassThrough)

    ZSql = ""
    ZSql = ZSql + "UPDATE CtaCtePrv SET "
    ZSql = ZSql + " Saldo = 0"
    ZSql = ZSql + " Where Saldo > -0.1 and Saldo < 0.1"
    spCtaCtePrv = ZSql
    Set rstCtaCtePrv = db.OpenRecordset(spCtaCtePrv, dbOpenSnapshot, dbSQLPassThrough)
    
    Call Cancelar_Click

End Sub

