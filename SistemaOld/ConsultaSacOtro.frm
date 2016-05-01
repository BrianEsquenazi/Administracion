VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgCosultaSacOtro 
   Caption         =   "Carga de SAC"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8370
   ScaleWidth      =   11775
   Begin VB.TextBox Numero 
      Alignment       =   1  'Right Justify
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
      Left            =   2760
      MaxLength       =   6
      TabIndex        =   170
      Text            =   " "
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Centro 
      Alignment       =   1  'Right Justify
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
      Left            =   5040
      MaxLength       =   6
      TabIndex        =   169
      Text            =   " "
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox Estado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6840
      TabIndex        =   168
      Top             =   480
      Width           =   2775
   End
   Begin VB.ComboBox Origen 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   167
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Emisor 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   166
      Text            =   " "
      Top             =   840
      Width           =   3615
   End
   Begin VB.TextBox Destino 
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
      Left            =   1320
      MaxLength       =   30
      TabIndex        =   165
      Text            =   " "
      Top             =   1200
      Width           =   3615
   End
   Begin VB.TextBox ResponsableEmisor 
      Alignment       =   1  'Right Justify
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
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   164
      Text            =   " "
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox ResponsableDestino 
      Alignment       =   1  'Right Justify
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
      Left            =   6840
      MaxLength       =   6
      TabIndex        =   163
      Text            =   " "
      Top             =   1200
      Width           =   855
   End
   Begin VB.TextBox Titulo 
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
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   162
      Text            =   " "
      Top             =   1560
      Width           =   9615
   End
   Begin VB.TextBox Ano 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      MaxLength       =   6
      TabIndex        =   161
      Text            =   " "
      Top             =   120
      Width           =   735
   End
   Begin Crystal.CrystalReport Listado 
      Left            =   11640
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab Tablas 
      Height          =   5535
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   9763
      _Version        =   327680
      Tabs            =   5
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Descripcion de la NC"
      TabPicture(0)   =   "ConsultaSacOtro.frx":0000
      Tab(0).ControlCount=   4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label19"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "IngresoCausa"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "IngresoNoCon"
      Tab(0).Control(3).Enabled=   0   'False
      TabCaption(1)   =   "Acciones"
      TabPicture(1)   =   "ConsultaSacOtro.frx":001C
      Tab(1).ControlCount=   33
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DesResponsable1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DesResponsable2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "DesResponsable3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "DesResponsable4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "DesResponsable5"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "DesResponsable6"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Plazo6"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Plazo5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Plazo4"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Plazo3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Plazo2"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Plazo1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Accion11"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "Accion12"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Accion21"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Accion22"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Accion31"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Accion32"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Accion41"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Accion42"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Accion51"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Accion52"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Accion61"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Accion62"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Responsable1"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Responsable2"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Responsable3"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Responsable4"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Responsable5"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Responsable6"
      Tab(1).Control(32).Enabled=   0   'False
      TabCaption(2)   =   "Implementacion"
      TabPicture(2)   =   "ConsultaSacOtro.frx":0038
      Tab(2).ControlCount=   53
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DesResponsable11"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label17"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label18"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label20"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DesResponsable12"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "DesResponsable13"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DesResponsable14"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "DesResponsable15"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DesResponsable16"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Label24"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Fecha6"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Fecha5"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Fecha4"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Fecha3"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Fecha2"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Fecha1"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Accion111"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Accion112"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Accion121"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Accion122"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Accion131"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Accion132"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Accion141"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Accion142"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Accion151"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Accion152"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Accion161"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Accion162"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Responsable11"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Comentario62"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Comentario61"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Comentario52"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Comentario51"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Comentario42"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Comentario41"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Comentario32"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Comentario31"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Comentario22"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Comentario21"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Comentario12"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Comentario11"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Responsable12"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Responsable13"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Responsable14"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Responsable16"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Responsable15"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Combo1"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Combo2"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Combo3"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Combo4"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Combo5"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Combo6"
      Tab(2).Control(52).Enabled=   0   'False
      TabCaption(3)   =   "Verificacion"
      TabPicture(3)   =   "ConsultaSacOtro.frx":0054
      Tab(3).ControlCount=   52
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label6"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label21"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DesResponsable21"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label23"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "DesResponsable22"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "DesResponsable23"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "DesResponsable24"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "DesResponsable25"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "DesResponsable26"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Fecha26"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Fecha25"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Fecha24"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Fecha23"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Fecha22"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Fecha21"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "Responsable21"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "Accion262"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Accion261"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Accion252"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "Accion251"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "Accion242"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "Accion241"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "Accion232"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "Accion231"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "Accion222"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "Accion221"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "Accion212"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "Accion211"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "Comentario262"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "Comentario261"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "Comentario252"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "Comentario251"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "Comentario242"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).Control(34)=   "Comentario241"
      Tab(3).Control(34).Enabled=   0   'False
      Tab(3).Control(35)=   "Comentario232"
      Tab(3).Control(35).Enabled=   0   'False
      Tab(3).Control(36)=   "Comentario231"
      Tab(3).Control(36).Enabled=   0   'False
      Tab(3).Control(37)=   "Comentario222"
      Tab(3).Control(37).Enabled=   0   'False
      Tab(3).Control(38)=   "Comentario221"
      Tab(3).Control(38).Enabled=   0   'False
      Tab(3).Control(39)=   "Comentario212"
      Tab(3).Control(39).Enabled=   0   'False
      Tab(3).Control(40)=   "Comentario211"
      Tab(3).Control(40).Enabled=   0   'False
      Tab(3).Control(41)=   "Responsable22"
      Tab(3).Control(41).Enabled=   0   'False
      Tab(3).Control(42)=   "Responsable23"
      Tab(3).Control(42).Enabled=   0   'False
      Tab(3).Control(43)=   "Responsable24"
      Tab(3).Control(43).Enabled=   0   'False
      Tab(3).Control(44)=   "Responsable25"
      Tab(3).Control(44).Enabled=   0   'False
      Tab(3).Control(45)=   "Responsable26"
      Tab(3).Control(45).Enabled=   0   'False
      Tab(3).Control(46)=   "Estado1"
      Tab(3).Control(46).Enabled=   0   'False
      Tab(3).Control(47)=   "Estado2"
      Tab(3).Control(47).Enabled=   0   'False
      Tab(3).Control(48)=   "Estado3"
      Tab(3).Control(48).Enabled=   0   'False
      Tab(3).Control(49)=   "Estado4"
      Tab(3).Control(49).Enabled=   0   'False
      Tab(3).Control(50)=   "Estado5"
      Tab(3).Control(50).Enabled=   0   'False
      Tab(3).Control(51)=   "Estado6"
      Tab(3).Control(51).Enabled=   0   'False
      TabCaption(4)   =   "Email"
      TabPicture(4)   =   "ConsultaSacOtro.frx":0070
      Tab(4).ControlCount=   30
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Boton1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Boton2"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Boton3"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Boton4"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Boton5"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Boton6"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Boton7"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Boton8"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Boton9"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Boton10"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "ImpreBoton1"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "ImpreBoton2"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "ImpreBoton3"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "ImpreBoton4"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "ImpreBoton5"
      Tab(4).Control(14).Enabled=   0   'False
      Tab(4).Control(15)=   "ImpreBoton6"
      Tab(4).Control(15).Enabled=   0   'False
      Tab(4).Control(16)=   "ImpreBoton7"
      Tab(4).Control(16).Enabled=   0   'False
      Tab(4).Control(17)=   "ImpreBoton8"
      Tab(4).Control(17).Enabled=   0   'False
      Tab(4).Control(18)=   "ImpreBoton10"
      Tab(4).Control(18).Enabled=   0   'False
      Tab(4).Control(19)=   "ImpreBoton9"
      Tab(4).Control(19).Enabled=   0   'False
      Tab(4).Control(20)=   "Email1"
      Tab(4).Control(20).Enabled=   0   'False
      Tab(4).Control(21)=   "Email2"
      Tab(4).Control(21).Enabled=   0   'False
      Tab(4).Control(22)=   "Email3"
      Tab(4).Control(22).Enabled=   0   'False
      Tab(4).Control(23)=   "Email4"
      Tab(4).Control(23).Enabled=   0   'False
      Tab(4).Control(24)=   "Email5"
      Tab(4).Control(24).Enabled=   0   'False
      Tab(4).Control(25)=   "Email8"
      Tab(4).Control(25).Enabled=   0   'False
      Tab(4).Control(26)=   "Email7"
      Tab(4).Control(26).Enabled=   0   'False
      Tab(4).Control(27)=   "Email6"
      Tab(4).Control(27).Enabled=   0   'False
      Tab(4).Control(28)=   "Email9"
      Tab(4).Control(28).Enabled=   0   'False
      Tab(4).Control(29)=   "Email10"
      Tab(4).Control(29).Enabled=   0   'False
      Begin VB.ComboBox Combo6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   192
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   191
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   190
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   189
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   188
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   187
         Top             =   4800
         Width           =   1335
      End
      Begin VB.TextBox IngresoNoCon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   184
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox IngresoCausa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4095
         Left            =   -69480
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   183
         Top             =   1200
         Width           =   5295
      End
      Begin VB.TextBox Email10 
         Height          =   1575
         Left            =   -66960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   150
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email9 
         Height          =   1575
         Left            =   -69000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   149
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email6 
         Height          =   1575
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   148
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email7 
         Height          =   1575
         Left            =   -72840
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   147
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email8 
         Height          =   1575
         Left            =   -70920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   146
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email5 
         Height          =   1575
         Left            =   -66960
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   145
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email4 
         Height          =   1575
         Left            =   -69000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   144
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email3 
         Height          =   1575
         Left            =   -70920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   143
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email2 
         Height          =   1575
         Left            =   -72840
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   142
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Email1 
         Height          =   1575
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   141
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Estado6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   140
         Top             =   4800
         Width           =   1335
      End
      Begin VB.ComboBox Estado5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   139
         Top             =   4080
         Width           =   1335
      End
      Begin VB.ComboBox Estado4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   138
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Estado3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   137
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox Estado2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   136
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox Estado1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -68880
         TabIndex        =   135
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Responsable26 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   132
         Text            =   " "
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Responsable25 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   129
         Text            =   " "
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Responsable24 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   126
         Text            =   " "
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Responsable23 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   123
         Text            =   " "
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Responsable22 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   120
         Text            =   " "
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Comentario211 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   119
         Text            =   " "
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Comentario212 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   118
         Text            =   " "
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox Comentario221 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   117
         Text            =   " "
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Comentario222 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   116
         Text            =   " "
         Top             =   2280
         Width           =   3615
      End
      Begin VB.TextBox Comentario231 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   115
         Text            =   " "
         Top             =   2640
         Width           =   3615
      End
      Begin VB.TextBox Comentario232 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   114
         Text            =   " "
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox Comentario241 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   113
         Text            =   " "
         Top             =   3360
         Width           =   3615
      End
      Begin VB.TextBox Comentario242 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   112
         Text            =   " "
         Top             =   3600
         Width           =   3615
      End
      Begin VB.TextBox Comentario251 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   111
         Text            =   " "
         Top             =   4080
         Width           =   3615
      End
      Begin VB.TextBox Comentario252 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   110
         Text            =   " "
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Comentario261 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   109
         Text            =   " "
         Top             =   4800
         Width           =   3615
      End
      Begin VB.TextBox Comentario262 
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
         Left            =   -67440
         MaxLength       =   50
         TabIndex        =   108
         Text            =   " "
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox Accion211 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   101
         Text            =   " "
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Accion212 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   100
         Text            =   " "
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Accion221 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   99
         Text            =   " "
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Accion222 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   98
         Text            =   " "
         Top             =   2280
         Width           =   3375
      End
      Begin VB.TextBox Accion231 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   97
         Text            =   " "
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Accion232 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   96
         Text            =   " "
         Top             =   2880
         Width           =   3375
      End
      Begin VB.TextBox Accion241 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   95
         Text            =   " "
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox Accion242 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   94
         Text            =   " "
         Top             =   3600
         Width           =   3375
      End
      Begin VB.TextBox Accion251 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   93
         Text            =   " "
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox Accion252 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   92
         Text            =   " "
         Top             =   4320
         Width           =   3375
      End
      Begin VB.TextBox Accion261 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   91
         Text            =   " "
         Top             =   4800
         Width           =   3375
      End
      Begin VB.TextBox Accion262 
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
         Left            =   -74760
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   90
         Text            =   " "
         Top             =   5040
         Width           =   3375
      End
      Begin VB.TextBox Responsable21 
         Alignment       =   1  'Right Justify
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
         Left            =   -71280
         MaxLength       =   6
         TabIndex        =   89
         Text            =   " "
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Responsable15 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   86
         Text            =   " "
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox Responsable16 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   83
         Text            =   " "
         Top             =   4800
         Width           =   615
      End
      Begin VB.TextBox Responsable14 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   80
         Text            =   " "
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox Responsable13 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   77
         Text            =   " "
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Responsable12 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   74
         Text            =   " "
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Comentario11 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   73
         Text            =   " "
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox Comentario12 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   72
         Text            =   " "
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox Comentario21 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   71
         Text            =   " "
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Comentario22 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   70
         Text            =   " "
         Top             =   2280
         Width           =   2895
      End
      Begin VB.TextBox Comentario31 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   69
         Text            =   " "
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox Comentario32 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   68
         Text            =   " "
         Top             =   2880
         Width           =   2895
      End
      Begin VB.TextBox Comentario41 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   67
         Text            =   " "
         Top             =   3360
         Width           =   2895
      End
      Begin VB.TextBox Comentario42 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   66
         Text            =   " "
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox Comentario51 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   65
         Text            =   " "
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox Comentario52 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   64
         Text            =   " "
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox Comentario61 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   63
         Text            =   " "
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox Comentario62 
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
         Left            =   8280
         MaxLength       =   50
         TabIndex        =   62
         Text            =   " "
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox Responsable11 
         Alignment       =   1  'Right Justify
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
         Left            =   3360
         MaxLength       =   6
         TabIndex        =   55
         Text            =   " "
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Accion162 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   54
         Text            =   " "
         Top             =   5040
         Width           =   3015
      End
      Begin VB.TextBox Accion161 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   53
         Text            =   " "
         Top             =   4800
         Width           =   3015
      End
      Begin VB.TextBox Accion152 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   52
         Text            =   " "
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox Accion151 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         Text            =   " "
         Top             =   4080
         Width           =   3015
      End
      Begin VB.TextBox Accion142 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Text            =   " "
         Top             =   3600
         Width           =   3015
      End
      Begin VB.TextBox Accion141 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Text            =   " "
         Top             =   3360
         Width           =   3015
      End
      Begin VB.TextBox Accion132 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   48
         Text            =   " "
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Accion131 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   47
         Text            =   " "
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox Accion122 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   46
         Text            =   " "
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Accion121 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   45
         Text            =   " "
         Top             =   2040
         Width           =   3015
      End
      Begin VB.TextBox Accion112 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   44
         Text            =   " "
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Accion111 
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
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   43
         Text            =   " "
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox Responsable6 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   40
         Text            =   " "
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox Responsable5 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   37
         Text            =   " "
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Responsable4 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   34
         Text            =   " "
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox Responsable3 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   31
         Text            =   " "
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Responsable2 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   28
         Text            =   " "
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Responsable1 
         Alignment       =   1  'Right Justify
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
         Left            =   -69360
         MaxLength       =   6
         TabIndex        =   22
         Text            =   " "
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Accion62 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   21
         Text            =   " "
         Top             =   5040
         Width           =   5295
      End
      Begin VB.TextBox Accion61 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   20
         Text            =   " "
         Top             =   4800
         Width           =   5295
      End
      Begin VB.TextBox Accion52 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   19
         Text            =   " "
         Top             =   4320
         Width           =   5295
      End
      Begin VB.TextBox Accion51 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   4080
         Width           =   5295
      End
      Begin VB.TextBox Accion42 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   17
         Text            =   " "
         Top             =   3600
         Width           =   5295
      End
      Begin VB.TextBox Accion41 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   3360
         Width           =   5295
      End
      Begin VB.TextBox Accion32 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   15
         Text            =   " "
         Top             =   2880
         Width           =   5295
      End
      Begin VB.TextBox Accion31 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   14
         Text            =   " "
         Top             =   2640
         Width           =   5295
      End
      Begin VB.TextBox Accion22 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   13
         Text            =   " "
         Top             =   2280
         Width           =   5295
      End
      Begin VB.TextBox Accion21 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   2040
         Width           =   5295
      End
      Begin VB.TextBox Accion12 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   11
         Text            =   " "
         Top             =   1560
         Width           =   5295
      End
      Begin VB.TextBox Accion11 
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
         Left            =   -74760
         MaxLength       =   50
         TabIndex        =   10
         Text            =   " "
         Top             =   1320
         Width           =   5295
      End
      Begin VB.TextBox WTexto2555 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   -360
         Width           =   375
      End
      Begin VB.TextBox WTexto1555 
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
         Left            =   1920
         TabIndex        =   8
         Top             =   -360
         Width           =   375
      End
      Begin MSMask.MaskEdBox WTexto3555 
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   -360
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         _Version        =   327680
         BackColor       =   16711935
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
      Begin MSMask.MaskEdBox Plazo1 
         Height          =   285
         Left            =   -65640
         TabIndex        =   25
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Plazo2 
         Height          =   285
         Left            =   -65640
         TabIndex        =   30
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Plazo3 
         Height          =   285
         Left            =   -65640
         TabIndex        =   33
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Plazo4 
         Height          =   285
         Left            =   -65640
         TabIndex        =   36
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Plazo5 
         Height          =   285
         Left            =   -65640
         TabIndex        =   39
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Plazo6 
         Height          =   285
         Left            =   -65640
         TabIndex        =   42
         Top             =   4800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha1 
         Height          =   285
         Left            =   5400
         TabIndex        =   58
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha2 
         Height          =   285
         Left            =   5400
         TabIndex        =   75
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha3 
         Height          =   285
         Left            =   5400
         TabIndex        =   78
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha4 
         Height          =   285
         Left            =   5400
         TabIndex        =   81
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha5 
         Height          =   285
         Left            =   5400
         TabIndex        =   84
         Top             =   4080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha6 
         Height          =   285
         Left            =   5400
         TabIndex        =   87
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha21 
         Height          =   285
         Left            =   -71280
         TabIndex        =   102
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha22 
         Height          =   285
         Left            =   -71280
         TabIndex        =   121
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha23 
         Height          =   285
         Left            =   -71280
         TabIndex        =   124
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha24 
         Height          =   285
         Left            =   -71280
         TabIndex        =   127
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha25 
         Height          =   285
         Left            =   -71280
         TabIndex        =   130
         Top             =   4320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin MSMask.MaskEdBox Fecha26 
         Height          =   285
         Left            =   -71280
         TabIndex        =   133
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   327680
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
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   6840
         TabIndex        =   193
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Causas que lo originaron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -69480
         TabIndex        =   186
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion de la No Conformidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -74880
         TabIndex        =   185
         Top             =   840
         Width           =   5295
      End
      Begin VB.Label ImpreBoton9 
         Alignment       =   2  'Center
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -67680
         TabIndex        =   160
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton10 
         Alignment       =   2  'Center
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -66840
         TabIndex        =   159
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton8 
         Alignment       =   2  'Center
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -68520
         TabIndex        =   158
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton7 
         Alignment       =   2  'Center
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69360
         TabIndex        =   157
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton6 
         Alignment       =   2  'Center
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70200
         TabIndex        =   156
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton5 
         Alignment       =   2  'Center
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71040
         TabIndex        =   155
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton4 
         Alignment       =   2  'Center
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71880
         TabIndex        =   154
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton3 
         Alignment       =   2  'Center
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72720
         TabIndex        =   153
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton2 
         Alignment       =   2  'Center
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -73560
         TabIndex        =   152
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label ImpreBoton1 
         Alignment       =   2  'Center
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74400
         TabIndex        =   151
         Top             =   5040
         Width           =   495
      End
      Begin VB.Image Boton10 
         Height          =   480
         Left            =   -66840
         MouseIcon       =   "ConsultaSacOtro.frx":008C
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":0396
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton9 
         Height          =   480
         Left            =   -67680
         MouseIcon       =   "ConsultaSacOtro.frx":04F9
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":0803
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton8 
         Height          =   480
         Left            =   -68520
         MouseIcon       =   "ConsultaSacOtro.frx":0966
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":0C70
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton7 
         Height          =   480
         Left            =   -69360
         MouseIcon       =   "ConsultaSacOtro.frx":0DD3
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":10DD
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton6 
         Height          =   480
         Left            =   -70200
         MouseIcon       =   "ConsultaSacOtro.frx":1240
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":154A
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton5 
         Height          =   480
         Left            =   -71040
         MouseIcon       =   "ConsultaSacOtro.frx":16AD
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":19B7
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton4 
         Height          =   480
         Left            =   -71880
         MouseIcon       =   "ConsultaSacOtro.frx":1B1A
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":1E24
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton3 
         Height          =   480
         Left            =   -72720
         MouseIcon       =   "ConsultaSacOtro.frx":1F87
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":2291
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton2 
         Height          =   480
         Left            =   -73560
         MouseIcon       =   "ConsultaSacOtro.frx":23F4
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":26FE
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Image Boton1 
         Height          =   480
         Left            =   -74400
         MouseIcon       =   "ConsultaSacOtro.frx":2861
         MousePointer    =   99  'Custom
         Picture         =   "ConsultaSacOtro.frx":2B6B
         ToolTipText     =   "Limpia la pantalla"
         Top             =   4560
         Width           =   480
      End
      Begin VB.Label DesResponsable26 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   134
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label DesResponsable25 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   131
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label DesResponsable24 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   128
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label DesResponsable23 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   125
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label DesResponsable22 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   122
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71280
         TabIndex        =   107
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label DesResponsable21 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -70560
         TabIndex        =   106
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -68880
         TabIndex        =   105
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74760
         TabIndex        =   104
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -67440
         TabIndex        =   103
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label DesResponsable16 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   88
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label DesResponsable15 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   85
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label DesResponsable14 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   82
         Top             =   3360
         Width           =   1215
      End
      Begin VB.Label DesResponsable13 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   79
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label DesResponsable12 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   76
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Comentarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8280
         TabIndex        =   61
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   60
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5400
         TabIndex        =   59
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label DesResponsable11 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   4080
         TabIndex        =   57
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         TabIndex        =   56
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label DesResponsable6 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   41
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label DesResponsable5 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   38
         Top             =   4080
         Width           =   2415
      End
      Begin VB.Label DesResponsable4 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   35
         Top             =   3360
         Width           =   2415
      End
      Begin VB.Label DesResponsable3 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   32
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label DesResponsable2 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   29
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acciones Correctivas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -74760
         TabIndex        =   27
         Top             =   960
         Width           =   5295
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Plazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -65640
         TabIndex        =   26
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label DesResponsable1 
         BackColor       =   &H00C0C000&
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
         Height          =   255
         Left            =   -68160
         TabIndex        =   24
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -69360
         TabIndex        =   23
         Top             =   960
         Width           =   3615
      End
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   9360
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   327680
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
   Begin VB.TextBox Ayuda 
      BackColor       =   &H00FFFFC0&
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
      Left            =   360
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Pantalla 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2460
      ItemData        =   "ConsultaSacOtro.frx":2CCE
      Left            =   360
      List            =   "ConsultaSacOtro.frx":2CD5
      TabIndex        =   6
      Top             =   4320
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.ListBox Opcion 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   1440
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label22 
      Caption         =   "Numero de SAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   182
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label15 
      Caption         =   "Centro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3960
      TabIndex        =   181
      Top             =   120
      Width           =   855
   End
   Begin VB.Label DesCentro 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   6000
      TabIndex        =   180
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label14 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5160
      TabIndex        =   179
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Emisor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   178
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   177
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "A Sector"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   176
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label DesResponsableEmisor 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   7800
      TabIndex        =   175
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label DesResponsableDestino 
      BackColor       =   &H00C0C000&
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
      Height          =   255
      Left            =   7800
      TabIndex        =   174
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5160
      TabIndex        =   173
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Responsable"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5160
      TabIndex        =   172
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Titulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   171
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Consulta 
      Height          =   480
      Left            =   4800
      MouseIcon       =   "ConsultaSacOtro.frx":2CE3
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaSacOtro.frx":2FED
      ToolTipText     =   "Consulta de Datos"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image CmdClose 
      Height          =   480
      Left            =   6720
      MouseIcon       =   "ConsultaSacOtro.frx":382F
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaSacOtro.frx":3B39
      ToolTipText     =   "Salida"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image CmdDelete 
      Height          =   480
      Left            =   8760
      MouseIcon       =   "ConsultaSacOtro.frx":437B
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaSacOtro.frx":4685
      ToolTipText     =   "Elimina el Registro"
      Top             =   7800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image CmdAdd 
      Height          =   480
      Left            =   3840
      MouseIcon       =   "ConsultaSacOtro.frx":4EC7
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaSacOtro.frx":51D1
      ToolTipText     =   "Graba los Datos Ingresados"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Image CmdLimpiar 
      Height          =   480
      Left            =   5760
      MouseIcon       =   "ConsultaSacOtro.frx":5A13
      MousePointer    =   99  'Custom
      Picture         =   "ConsultaSacOtro.frx":5D1D
      ToolTipText     =   "Limpia la pantalla"
      Top             =   7800
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   8400
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "PrgCosultaSacOtro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstCargaSac As Recordset
Dim spCargaSac As String
Dim rstCentroSac As Recordset
Dim spCentroSac As String
Dim rstResponsableSac As Recordset
Dim spResponsableSac As String

Dim XParam As String
Dim ZZLugar As Integer

Sub Imprime_Descripcion()
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable21.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable21.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable22.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable22.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable23.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable23.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable24.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable24.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable25.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable25.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM ResponsableSac"
    Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable26.Text + "'"
    spResponsableSac = Sql1 + Sql2 + Sql3
    Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstResponsableSac.RecordCount > 0 Then
        DesResponsable26.Caption = Trim(rstResponsableSac!Descripcion)
        rstResponsableSac.Close
    End If
    
    Sql1 = "Select *"
    Sql2 = " FROM CentroSac"
    Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
    spCentroSac = Sql1 + Sql2 + Sql3
    Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCentroSac.RecordCount > 0 Then
        DesCentro.Caption = Trim(rstCentroSac!Descripcion)
        rstCentroSac.Close
    End If

End Sub


Sub Verifica_datos()
End Sub

Sub Imprime_Datos()

    On Error GoTo WError

    ZCentro = Centro.Text
    ZNumero = Numero.Text
    
    Call CmdLimpiar_Click
    
    ZExiste = "N"
    
    Centro.Text = ZCentro
    Numero.Text = ZNumero
        
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM CargaSac"
    ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
    ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
    spCargaSac = ZSql
    Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
    If rstCargaSac.RecordCount > 0 Then
        Fecha.Text = rstCargaSac!Fecha
        Seccion.Text = Str$(rstCargaSac!Seccion)
        Estado.Text = Str$(rstCargaSac!Estado)
        Origen.Text = Trim(rstCargaSac!Origen)
        Motivo.Text = Trim(rstCargaSac!Motivo)
        
        DescripcionI.Text = Trim(rstCargaSac!DescripcionI)
        DescripcionII.Text = Trim(rstCargaSac!DescripcionII)
        DescripcionIII.Text = Trim(rstCargaSac!DescripcionIII)
        DescripcionIV.Text = Trim(rstCargaSac!DescripcionIV)
        DescripcionV.Text = Trim(rstCargaSac!DescripcionV)
        DescripcionVI.Text = Trim(rstCargaSac!DescripcionVI)
        DescripcionVII.Text = Trim(rstCargaSac!DescripcionVII)
        DescripcionVIII.Text = Trim(rstCargaSac!DescripcionVIII)
        DescripcionIX.Text = Trim(rstCargaSac!DescripcionIX)
        DescripcionX.Text = Trim(rstCargaSac!DescripcionX)
        
        CausaI.Text = Trim(rstCargaSac!CausaI)
        CausaII.Text = Trim(rstCargaSac!CausaII)
        CausaIII.Text = Trim(rstCargaSac!CausaIII)
        CausaIV.Text = Trim(rstCargaSac!CausaIV)
        CausaV.Text = Trim(rstCargaSac!CausaV)
        CausaVI.Text = Trim(rstCargaSac!CausaVI)
        CausaVII.Text = Trim(rstCargaSac!CausaVII)
        CausaVIII.Text = Trim(rstCargaSac!CausaVIII)
        CausaIX.Text = Trim(rstCargaSac!CausaIX)
        CausaX.Text = Trim(rstCargaSac!CausaX)
        
        rstCargaSac.Close
    End If
    Exit Sub
    
WError:
    Resume Next
    
End Sub

Private Sub Boton1_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email1.Height = 3735
    Email1.Left = 240
    Email1.Top = 840
    Email1.Width = 10695
    Email1.Visible = True
    
    ImpreBoton1.Visible = True
    
End Sub

Private Sub Boton2_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email2.Height = 3735
    Email2.Left = 240
    Email2.Top = 840
    Email2.Width = 10695
    Email2.Visible = True
    
    ImpreBoton2.Visible = True
    
End Sub

Private Sub Boton3_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email3.Height = 3735
    Email3.Left = 240
    Email3.Top = 840
    Email3.Width = 10695
    Email3.Visible = True
    
    ImpreBoton3.Visible = True
    
End Sub


Private Sub Boton4_Click()


    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email4.Height = 3735
    Email4.Left = 240
    Email4.Top = 840
    Email4.Width = 10695
    Email4.Visible = True
    
    ImpreBoton4.Visible = True

End Sub

Private Sub Boton5_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email5.Height = 3735
    Email5.Left = 240
    Email5.Top = 840
    Email5.Width = 10695
    Email5.Visible = True
    
    ImpreBoton5.Visible = True
    

End Sub

Private Sub Boton6_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email6.Height = 3735
    Email6.Left = 240
    Email6.Top = 840
    Email6.Width = 10695
    Email6.Visible = True
    
    ImpreBoton6.Visible = True
    

End Sub

Private Sub Boton7_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email7.Height = 3735
    Email7.Left = 240
    Email7.Top = 840
    Email7.Width = 10695
    Email7.Visible = True
    
    ImpreBoton7.Visible = True
    

End Sub

Private Sub Boton8_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email8.Height = 3735
    Email8.Left = 240
    Email8.Top = 840
    Email8.Width = 10695
    Email8.Visible = True
    
    ImpreBoton8.Visible = True
    

End Sub

Private Sub Boton9_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email9.Height = 3735
    Email9.Left = 240
    Email9.Top = 840
    Email9.Width = 10695
    Email9.Visible = True
    
    ImpreBoton9.Visible = True
    

End Sub

Private Sub Boton10_Click()

    Email1.Visible = False
    Email2.Visible = False
    Email3.Visible = False
    Email4.Visible = False
    Email5.Visible = False
    Email6.Visible = False
    Email7.Visible = False
    Email8.Visible = False
    Email9.Visible = False
    Email10.Visible = False
    
    ImpreBoton1.Visible = False
    ImpreBoton2.Visible = False
    ImpreBoton3.Visible = False
    ImpreBoton4.Visible = False
    ImpreBoton5.Visible = False
    ImpreBoton6.Visible = False
    ImpreBoton7.Visible = False
    ImpreBoton8.Visible = False
    ImpreBoton9.Visible = False
    ImpreBoton10.Visible = False
    
    Email10.Height = 3735
    Email10.Left = 240
    Email10.Top = 840
    Email10.Width = 10695
    Email10.Visible = True
    
    ImpreBoton10.Visible = True
    

End Sub

Private Sub cmdAdd_Click()

    If Centro.Text <> "" And Numero.Text <> "" Then
    
        WOrdFecha = Right$(Fecha.Text, 4) + Mid$(Fecha.Text, 4, 2) + Left$(Fecha.Text, 2)
        Auxi1 = Centro.Text
        Auxi2 = Numero.Text
        Call Ceros(Auxi1, 4)
        Call Ceros(Auxi2, 6)
        WClave = Auxi1 + Auxi2
    
        ZSql = ""
        ZSql = ZSql + "Select Clave, Centro, Numero"
        ZSql = ZSql + " FROM CargaSac"
        ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
        ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            rstCargaSac.Close
            
            ZSql = ""
            ZSql = ZSql + "UPDATE CargaSac SET "
            ZSql = ZSql + " Fecha = " + "'" + Fecha.Text + "',"
            ZSql = ZSql + " OrdFecha = " + "'" + WOrdFecha + "',"
            ZSql = ZSql + " Seccion = " + "'" + Str$(Seccion.ListIndex) + "',"
            ZSql = ZSql + " Estado = " + "'" + Str$(Estado.ListIndex) + "',"
            ZSql = ZSql + " Origen = " + "'" + Origen.Text + "',"
            ZSql = ZSql + " Motivo = " + "'" + Motivo.Text + "',"
            ZSql = ZSql + " Descripcion1 = " + "'" + Descripcion1.Text + "',"
            ZSql = ZSql + " Descripcion2 = " + "'" + Descripcion2.Text + "',"
            ZSql = ZSql + " Descripcion3 = " + "'" + Descripcion3.Text + "',"
            ZSql = ZSql + " Descripcion4 = " + "'" + Descripcion4.Text + "',"
            ZSql = ZSql + " Descripcion5 = " + "'" + Descripcion5.Text + "',"
            ZSql = ZSql + " Descripcion6 = " + "'" + Descripcion6.Text + "',"
            ZSql = ZSql + " Descripcion7 = " + "'" + Descripcion7.Text + "',"
            ZSql = ZSql + " Descripcion8 = " + "'" + Descripcion8.Text + "',"
            ZSql = ZSql + " Descripcion9 = " + "'" + Descripcion9.Text + "',"
            ZSql = ZSql + " Descripcion10 = " + "'" + Descripcion10.Text + "',"
            ZSql = ZSql + " Causa1 = " + "'" + Causa1.Text + "',"
            ZSql = ZSql + " Causa2 = " + "'" + Causa2.Text + "',"
            ZSql = ZSql + " Causa3 = " + "'" + Causa3.Text + "',"
            ZSql = ZSql + " Causa4 = " + "'" + Causa4.Text + "',"
            ZSql = ZSql + " Causa5 = " + "'" + Causa5.Text + "',"
            ZSql = ZSql + " Causa6 = " + "'" + Causa6.Text + "',"
            ZSql = ZSql + " Causa7 = " + "'" + Causa7.Text + "',"
            ZSql = ZSql + " Causa8 = " + "'" + Causa8.Text + "',"
            ZSql = ZSql + " Causa9 = " + "'" + Causa9.Text + "',"
            ZSql = ZSql + " Causa10 = " + "'" + Causa10.Text + "',"
            ZSql = ZSql + " Accion11 = " + "'" + Accion11.Text + "',"
            ZSql = ZSql + " Accion12 = " + "'" + Accion12.Text + "',"
            ZSql = ZSql + " Accion21 = " + "'" + Accion21.Text + "',"
            ZSql = ZSql + " Accion22 = " + "'" + Accion22.Text + "',"
            ZSql = ZSql + " Accion31 = " + "'" + Accion31.Text + "',"
            ZSql = ZSql + " Accion32 = " + "'" + Accion32.Text + "',"
            ZSql = ZSql + " Accion41 = " + "'" + Accion41.Text + "',"
            ZSql = ZSql + " Accion42 = " + "'" + Accion42.Text + "',"
            ZSql = ZSql + " Accion51 = " + "'" + Accion51.Text + "',"
            ZSql = ZSql + " Accion52 = " + "'" + Accion52.Text + "',"
            ZSql = ZSql + " Accion61 = " + "'" + Accion61.Text + "',"
            ZSql = ZSql + " Accion62 = " + "'" + Accion62.Text + "',"
            ZSql = ZSql + " Responsable1 = " + "'" + Responsable1.Text + "',"
            ZSql = ZSql + " Responsable2 = " + "'" + Responsable2.Text + "',"
            ZSql = ZSql + " Responsable3 = " + "'" + Responsable3.Text + "',"
            ZSql = ZSql + " Responsable4 = " + "'" + Responsable4.Text + "',"
            ZSql = ZSql + " Responsable5 = " + "'" + Responsable5.Text + "',"
            ZSql = ZSql + " Responsable6 = " + "'" + Responsable6.Text + "',"
            ZSql = ZSql + " Plazo1 = " + "'" + Plazo1.Text + "',"
            ZSql = ZSql + " Plazo2 = " + "'" + Plazo2.Text + "',"
            ZSql = ZSql + " Plazo3 = " + "'" + Plazo3.Text + "',"
            ZSql = ZSql + " Plazo4 = " + "'" + Plazo4.Text + "',"
            ZSql = ZSql + " Plazo5 = " + "'" + Plazo5.Text + "',"
            ZSql = ZSql + " Plazo6 = " + "'" + Plazo6.Text + "',"
            ZSql = ZSql + " Responsable11 = " + "'" + Responsable11.Text + "',"
            ZSql = ZSql + " Responsable12 = " + "'" + Responsable12.Text + "',"
            ZSql = ZSql + " Responsable13 = " + "'" + Responsable13.Text + "',"
            ZSql = ZSql + " Responsable14 = " + "'" + Responsable14.Text + "',"
            ZSql = ZSql + " Responsable15 = " + "'" + Responsable15.Text + "',"
            ZSql = ZSql + " Responsable16 = " + "'" + Responsable16.Text + "',"
            ZSql = ZSql + " Fecha1 = " + "'" + Fecha1.Text + "',"
            ZSql = ZSql + " Fecha2 = " + "'" + Fecha2.Text + "',"
            ZSql = ZSql + " Fecha3 = " + "'" + Fecha3.Text + "',"
            ZSql = ZSql + " Fecha4 = " + "'" + Fecha4.Text + "',"
            ZSql = ZSql + " Fecha5 = " + "'" + Fecha5.Text + "',"
            ZSql = ZSql + " Fecha6 = " + "'" + Fecha6.Text + "',"
            ZSql = ZSql + " Comentario11 = " + "'" + Comentario11.Text + "',"
            ZSql = ZSql + " Comentario12 = " + "'" + Comentario12.Text + "',"
            ZSql = ZSql + " Comentario21 = " + "'" + Comentario21.Text + "',"
            ZSql = ZSql + " Comentario22 = " + "'" + Comentario22.Text + "',"
            ZSql = ZSql + " Comentario31 = " + "'" + Comentario31.Text + "',"
            ZSql = ZSql + " Comentario32 = " + "'" + Comentario32.Text + "',"
            ZSql = ZSql + " Comentario41 = " + "'" + Comentario41.Text + "',"
            ZSql = ZSql + " Comentario42 = " + "'" + Comentario42.Text + "',"
            ZSql = ZSql + " Comentario51 = " + "'" + Comentario51.Text + "',"
            ZSql = ZSql + " Comentario52 = " + "'" + Comentario52.Text + "',"
            ZSql = ZSql + " Comentario61 = " + "'" + Comentario61.Text + "',"
            ZSql = ZSql + " Comentario62 = " + "'" + Comentario62.Text + "',"
            ZSql = ZSql + " Responsable21 = " + "'" + Responsable21.Text + "',"
            ZSql = ZSql + " Responsable22 = " + "'" + Responsable22.Text + "',"
            ZSql = ZSql + " Responsable23 = " + "'" + Responsable23.Text + "',"
            ZSql = ZSql + " Responsable24 = " + "'" + Responsable24.Text + "',"
            ZSql = ZSql + " Responsable25 = " + "'" + Responsable25.Text + "',"
            ZSql = ZSql + " Responsable26 = " + "'" + Responsable26.Text + "',"
            ZSql = ZSql + " Fecha21 = " + "'" + Fecha21.Text + "',"
            ZSql = ZSql + " Fecha22 = " + "'" + Fecha22.Text + "',"
            ZSql = ZSql + " Fecha23 = " + "'" + Fecha23.Text + "',"
            ZSql = ZSql + " Fecha24 = " + "'" + Fecha24.Text + "',"
            ZSql = ZSql + " Fecha25 = " + "'" + Fecha25.Text + "',"
            ZSql = ZSql + " Fecha26 = " + "'" + Fecha26.Text + "',"
            ZSql = ZSql + " Estado1 = " + "'" + Str$(Estado1.ListIndex) + "',"
            ZSql = ZSql + " Estado2 = " + "'" + Str$(Estado2.ListIndex) + "',"
            ZSql = ZSql + " Estado3 = " + "'" + Str$(Estado3.ListIndex) + "',"
            ZSql = ZSql + " Estado4 = " + "'" + Str$(Estado4.ListIndex) + "',"
            ZSql = ZSql + " Estado5 = " + "'" + Str$(Estado5.ListIndex) + "',"
            ZSql = ZSql + " Estado6 = " + "'" + Str$(Estado6.ListIndex) + "',"
            ZSql = ZSql + " Comentario211 = " + "'" + Comentario211.Text + "',"
            ZSql = ZSql + " Comentario212 = " + "'" + Comentario212.Text + "',"
            ZSql = ZSql + " Comentario221 = " + "'" + Comentario221.Text + "',"
            ZSql = ZSql + " Comentario222 = " + "'" + Comentario222.Text + "',"
            ZSql = ZSql + " Comentario231 = " + "'" + Comentario231.Text + "',"
            ZSql = ZSql + " Comentario232 = " + "'" + Comentario232.Text + "',"
            ZSql = ZSql + " Comentario241 = " + "'" + Comentario241.Text + "',"
            ZSql = ZSql + " Comentario242 = " + "'" + Comentario242.Text + "',"
            ZSql = ZSql + " Comentario251 = " + "'" + Comentario251.Text + "',"
            ZSql = ZSql + " Comentario252 = " + "'" + Comentario252.Text + "',"
            ZSql = ZSql + " Comentario261 = " + "'" + Comentario261.Text + "',"
            ZSql = ZSql + " Comentario262 = " + "'" + Comentario262.Text + "',"
            ZSql = ZSql + " Email1 = " + "'" + Email1.Text + "',"
            ZSql = ZSql + " Email2 = " + "'" + Email2.Text + "',"
            ZSql = ZSql + " Email3 = " + "'" + Email3.Text + "',"
            ZSql = ZSql + " Email4 = " + "'" + Email4.Text + "',"
            ZSql = ZSql + " Email5 = " + "'" + Email5.Text + "',"
            ZSql = ZSql + " Email6 = " + "'" + Email6.Text + "',"
            ZSql = ZSql + " Email7 = " + "'" + Email7.Text + "',"
            ZSql = ZSql + " Email8 = " + "'" + Email8.Text + "',"
            ZSql = ZSql + " Email9 = " + "'" + Email9.Text + "',"
            ZSql = ZSql + " Email10 = " + "'" + Email10.Text + "',"
            ZSql = ZSql + " DesResponsable1 = " + "'" + DesResponsable1.Caption + "',"
            ZSql = ZSql + " DesResponsable2 = " + "'" + DesResponsable2.Caption + "',"
            ZSql = ZSql + " DesResponsable3 = " + "'" + DesResponsable3.Caption + "',"
            ZSql = ZSql + " DesResponsable4 = " + "'" + DesResponsable4.Caption + "',"
            ZSql = ZSql + " DesResponsable5 = " + "'" + DesResponsable5.Caption + "',"
            ZSql = ZSql + " DesResponsable6 = " + "'" + DesResponsable6.Caption + "',"
            ZSql = ZSql + " DesResponsable7 = " + "'" + DesResponsable11.Caption + "',"
            ZSql = ZSql + " DesResponsable8 = " + "'" + DesResponsable12.Caption + "',"
            ZSql = ZSql + " DesResponsable9 = " + "'" + DesResponsable13.Caption + "',"
            ZSql = ZSql + " DesResponsable10 = " + "'" + DesResponsable14.Caption + "',"
            ZSql = ZSql + " DesResponsable11 = " + "'" + DesResponsable15.Caption + "',"
            ZSql = ZSql + " DesResponsable12 = " + "'" + DesResponsable16.Caption + "',"
            ZSql = ZSql + " DesResponsable13 = " + "'" + DesResponsable21.Caption + "',"
            ZSql = ZSql + " DesResponsable14 = " + "'" + DesResponsable22.Caption + "',"
            ZSql = ZSql + " DesResponsable15 = " + "'" + DesResponsable23.Caption + "',"
            ZSql = ZSql + " DesResponsable16 = " + "'" + DesResponsable24.Caption + "',"
            ZSql = ZSql + " DesResponsable17 = " + "'" + DesResponsable25.Caption + "',"
            ZSql = ZSql + " DesResponsable18 = " + "'" + DesResponsable26.Caption + "',"
            ZSql = ZSql + " DesSeccion = " + "'" + Seccion.Text + "',"
            ZSql = ZSql + " DesEstado = " + "'" + Estado.Text + "',"
            ZSql = ZSql + " DesEstado1 = " + "'" + Estado1.Text + "',"
            ZSql = ZSql + " DesEstado2 = " + "'" + Estado2.Text + "',"
            ZSql = ZSql + " DesEstado3 = " + "'" + Estado3.Text + "',"
            ZSql = ZSql + " DesEstado4 = " + "'" + Estado4.Text + "',"
            ZSql = ZSql + " DesEstado5 = " + "'" + Estado5.Text + "',"
            ZSql = ZSql + " DesEstado6 = " + "'" + Estado6.Text + "',"
            ZSql = ZSql + " DesCentro = " + "'" + DesCentro.Caption + "'"
            ZSql = ZSql + " Where Centro = " + "'" + Centro.Text + "'"
            ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
                Else
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CargaSac ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Centro ,"
            ZSql = ZSql + "Numero ,"
            ZSql = ZSql + "Fecha ,"
            ZSql = ZSql + "OrdFecha ,"
            ZSql = ZSql + "Seccion ,"
            ZSql = ZSql + "Estado ,"
            ZSql = ZSql + "Origen ,"
            ZSql = ZSql + "Motivo ,"
            ZSql = ZSql + "Descripcion1 ,"
            ZSql = ZSql + "Descripcion2 ,"
            ZSql = ZSql + "Descripcion3 ,"
            ZSql = ZSql + "Descripcion4 ,"
            ZSql = ZSql + "Descripcion5 ,"
            ZSql = ZSql + "Descripcion6 ,"
            ZSql = ZSql + "Descripcion7 ,"
            ZSql = ZSql + "Descripcion8 ,"
            ZSql = ZSql + "Descripcion9 ,"
            ZSql = ZSql + "Descripcion10 ,"
            ZSql = ZSql + "Causa1 ,"
            ZSql = ZSql + "Causa2 ,"
            ZSql = ZSql + "Causa3 ,"
            ZSql = ZSql + "Causa4 ,"
            ZSql = ZSql + "Causa5 ,"
            ZSql = ZSql + "Causa6 ,"
            ZSql = ZSql + "Causa7 ,"
            ZSql = ZSql + "Causa8 ,"
            ZSql = ZSql + "Causa9 ,"
            ZSql = ZSql + "Causa10 ,"
            ZSql = ZSql + "Accion11 ,"
            ZSql = ZSql + "Accion12 ,"
            ZSql = ZSql + "Accion21 ,"
            ZSql = ZSql + "Accion22 ,"
            ZSql = ZSql + "Accion31 ,"
            ZSql = ZSql + "Accion32 ,"
            ZSql = ZSql + "Accion41 ,"
            ZSql = ZSql + "Accion42 ,"
            ZSql = ZSql + "Accion51 ,"
            ZSql = ZSql + "Accion52 ,"
            ZSql = ZSql + "Accion61 ,"
            ZSql = ZSql + "Accion62 ,"
            ZSql = ZSql + "Responsable1 ,"
            ZSql = ZSql + "Responsable2 ,"
            ZSql = ZSql + "Responsable3 ,"
            ZSql = ZSql + "Responsable4 ,"
            ZSql = ZSql + "Responsable5 ,"
            ZSql = ZSql + "Responsable6 ,"
            ZSql = ZSql + "Plazo1 ,"
            ZSql = ZSql + "Plazo2 ,"
            ZSql = ZSql + "Plazo3 ,"
            ZSql = ZSql + "Plazo4 ,"
            ZSql = ZSql + "Plazo5 ,"
            ZSql = ZSql + "Plazo6 ,"
            ZSql = ZSql + "Responsable11 ,"
            ZSql = ZSql + "Responsable12 ,"
            ZSql = ZSql + "Responsable13 ,"
            ZSql = ZSql + "Responsable14 ,"
            ZSql = ZSql + "Responsable15 ,"
            ZSql = ZSql + "Responsable16 ,"
            ZSql = ZSql + "Fecha1 ,"
            ZSql = ZSql + "Fecha2 ,"
            ZSql = ZSql + "Fecha3 ,"
            ZSql = ZSql + "Fecha4 ,"
            ZSql = ZSql + "Fecha5 ,"
            ZSql = ZSql + "Fecha6 ,"
            ZSql = ZSql + "Comentario11 ,"
            ZSql = ZSql + "Comentario12 ,"
            ZSql = ZSql + "Comentario21 ,"
            ZSql = ZSql + "Comentario22 ,"
            ZSql = ZSql + "Comentario31 ,"
            ZSql = ZSql + "Comentario32 ,"
            ZSql = ZSql + "Comentario41 ,"
            ZSql = ZSql + "Comentario42 ,"
            ZSql = ZSql + "Comentario51 ,"
            ZSql = ZSql + "Comentario52 ,"
            ZSql = ZSql + "Comentario61 ,"
            ZSql = ZSql + "Comentario62 ,"
            ZSql = ZSql + "Responsable21 ,"
            ZSql = ZSql + "Responsable22 ,"
            ZSql = ZSql + "Responsable23 ,"
            ZSql = ZSql + "Responsable24 ,"
            ZSql = ZSql + "Responsable25 ,"
            ZSql = ZSql + "Responsable26 ,"
            ZSql = ZSql + "Fecha21 ,"
            ZSql = ZSql + "Fecha22 ,"
            ZSql = ZSql + "Fecha23 ,"
            ZSql = ZSql + "Fecha24 ,"
            ZSql = ZSql + "Fecha25 ,"
            ZSql = ZSql + "Fecha26 ,"
            ZSql = ZSql + "Estado1 ,"
            ZSql = ZSql + "Estado2 ,"
            ZSql = ZSql + "Estado3 ,"
            ZSql = ZSql + "Estado4 ,"
            ZSql = ZSql + "Estado5 ,"
            ZSql = ZSql + "Estado6 ,"
            ZSql = ZSql + "Comentario211 ,"
            ZSql = ZSql + "Comentario212 ,"
            ZSql = ZSql + "Comentario221 ,"
            ZSql = ZSql + "Comentario222 ,"
            ZSql = ZSql + "Comentario231 ,"
            ZSql = ZSql + "Comentario232 ,"
            ZSql = ZSql + "Comentario241 ,"
            ZSql = ZSql + "Comentario242 ,"
            ZSql = ZSql + "Comentario251 ,"
            ZSql = ZSql + "Comentario252 ,"
            ZSql = ZSql + "Comentario261 ,"
            ZSql = ZSql + "Comentario262 ,"
            ZSql = ZSql + "Email1 ,"
            ZSql = ZSql + "Email2 ,"
            ZSql = ZSql + "Email3 ,"
            ZSql = ZSql + "Email4 ,"
            ZSql = ZSql + "Email5 ,"
            ZSql = ZSql + "Email6 ,"
            ZSql = ZSql + "Email7 ,"
            ZSql = ZSql + "Email8 ,"
            ZSql = ZSql + "Email9 ,"
            ZSql = ZSql + "Email10 ,"
            ZSql = ZSql + "DesResponsable1 ,"
            ZSql = ZSql + "DesResponsable2 ,"
            ZSql = ZSql + "DesResponsable3 ,"
            ZSql = ZSql + "DesResponsable4 ,"
            ZSql = ZSql + "DesResponsable5 ,"
            ZSql = ZSql + "DesResponsable6 ,"
            ZSql = ZSql + "DesResponsable7 ,"
            ZSql = ZSql + "DesResponsable8 ,"
            ZSql = ZSql + "DesResponsable9 ,"
            ZSql = ZSql + "DesResponsable10 ,"
            ZSql = ZSql + "DesResponsable11 ,"
            ZSql = ZSql + "DesResponsable12 ,"
            ZSql = ZSql + "DesResponsable13 ,"
            ZSql = ZSql + "DesResponsable14 ,"
            ZSql = ZSql + "DesResponsable15 ,"
            ZSql = ZSql + "DesResponsable16 ,"
            ZSql = ZSql + "DesResponsable17 ,"
            ZSql = ZSql + "DesResponsable18 ,"
            ZSql = ZSql + "DesSeccion ,"
            ZSql = ZSql + "DesEstado ,"
            ZSql = ZSql + "DesEstado1 ,"
            ZSql = ZSql + "DesEstado2 ,"
            ZSql = ZSql + "DesEstado3 ,"
            ZSql = ZSql + "DesEstado4 ,"
            ZSql = ZSql + "DesEstado5 ,"
            ZSql = ZSql + "DesEstado6 ,"
            ZSql = ZSql + "DesCentro )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WClave + "',"
            ZSql = ZSql + "'" + Centro.Text + "',"
            ZSql = ZSql + "'" + Numero.Text + "',"
            ZSql = ZSql + "'" + Fecha.Text + "',"
            ZSql = ZSql + "'" + WOrdFecha + "',"
            ZSql = ZSql + "'" + Str$(Seccion.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado.ListIndex) + "',"
            ZSql = ZSql + "'" + Origen.Text + "',"
            ZSql = ZSql + "'" + Motivo.Text + "',"
            ZSql = ZSql + "'" + Descripcion1.Text + "',"
            ZSql = ZSql + "'" + Descripcion2.Text + "',"
            ZSql = ZSql + "'" + Descripcion3.Text + "',"
            ZSql = ZSql + "'" + Descripcion4.Text + "',"
            ZSql = ZSql + "'" + Descripcion5.Text + "',"
            ZSql = ZSql + "'" + Descripcion6.Text + "',"
            ZSql = ZSql + "'" + Descripcion7.Text + "',"
            ZSql = ZSql + "'" + Descripcion8.Text + "',"
            ZSql = ZSql + "'" + Descripcion9.Text + "',"
            ZSql = ZSql + "'" + Descripcion10.Text + "',"
            ZSql = ZSql + "'" + Causa1.Text + "',"
            ZSql = ZSql + "'" + Causa2.Text + "',"
            ZSql = ZSql + "'" + Causa3.Text + "',"
            ZSql = ZSql + "'" + Causa4.Text + "',"
            ZSql = ZSql + "'" + Causa5.Text + "',"
            ZSql = ZSql + "'" + Causa6.Text + "',"
            ZSql = ZSql + "'" + Causa7.Text + "',"
            ZSql = ZSql + "'" + Causa8.Text + "',"
            ZSql = ZSql + "'" + Causa9.Text + "',"
            ZSql = ZSql + "'" + Causa10.Text + "',"
            ZSql = ZSql + "'" + Accion11.Text + "',"
            ZSql = ZSql + "'" + Accion12.Text + "',"
            ZSql = ZSql + "'" + Accion21.Text + "',"
            ZSql = ZSql + "'" + Accion22.Text + "',"
            ZSql = ZSql + "'" + Accion31.Text + "',"
            ZSql = ZSql + "'" + Accion32.Text + "',"
            ZSql = ZSql + "'" + Accion41.Text + "',"
            ZSql = ZSql + "'" + Accion42.Text + "',"
            ZSql = ZSql + "'" + Accion51.Text + "',"
            ZSql = ZSql + "'" + Accion52.Text + "',"
            ZSql = ZSql + "'" + Accion61.Text + "',"
            ZSql = ZSql + "'" + Accion62.Text + "',"
            ZSql = ZSql + "'" + Responsable1.Text + "',"
            ZSql = ZSql + "'" + Responsable2.Text + "',"
            ZSql = ZSql + "'" + Responsable3.Text + "',"
            ZSql = ZSql + "'" + Responsable4.Text + "',"
            ZSql = ZSql + "'" + Responsable5.Text + "',"
            ZSql = ZSql + "'" + Responsable6.Text + "',"
            ZSql = ZSql + "'" + Plazo1.Text + "',"
            ZSql = ZSql + "'" + Plazo2.Text + "',"
            ZSql = ZSql + "'" + Plazo3.Text + "',"
            ZSql = ZSql + "'" + Plazo4.Text + "',"
            ZSql = ZSql + "'" + Plazo5.Text + "',"
            ZSql = ZSql + "'" + Plazo6.Text + "',"
            ZSql = ZSql + "'" + Responsable11.Text + "',"
            ZSql = ZSql + "'" + Responsable12.Text + "',"
            ZSql = ZSql + "'" + Responsable13.Text + "',"
            ZSql = ZSql + "'" + Responsable14.Text + "',"
            ZSql = ZSql + "'" + Responsable15.Text + "',"
            ZSql = ZSql + "'" + Responsable16.Text + "',"
            ZSql = ZSql + "'" + Fecha1.Text + "',"
            ZSql = ZSql + "'" + Fecha2.Text + "',"
            ZSql = ZSql + "'" + Fecha3.Text + "',"
            ZSql = ZSql + "'" + Fecha4.Text + "',"
            ZSql = ZSql + "'" + Fecha5.Text + "',"
            ZSql = ZSql + "'" + Fecha6.Text + "',"
            ZSql = ZSql + "'" + Comentario11.Text + "',"
            ZSql = ZSql + "'" + Comentario12.Text + "',"
            ZSql = ZSql + "'" + Comentario21.Text + "',"
            ZSql = ZSql + "'" + Comentario22.Text + "',"
            ZSql = ZSql + "'" + Comentario31.Text + "',"
            ZSql = ZSql + "'" + Comentario32.Text + "',"
            ZSql = ZSql + "'" + Comentario41.Text + "',"
            ZSql = ZSql + "'" + Comentario42.Text + "',"
            ZSql = ZSql + "'" + Comentario51.Text + "',"
            ZSql = ZSql + "'" + Comentario52.Text + "',"
            ZSql = ZSql + "'" + Comentario61.Text + "',"
            ZSql = ZSql + "'" + Comentario62.Text + "',"
            ZSql = ZSql + "'" + Responsable21.Text + "',"
            ZSql = ZSql + "'" + Responsable22.Text + "',"
            ZSql = ZSql + "'" + Responsable23.Text + "',"
            ZSql = ZSql + "'" + Responsable24.Text + "',"
            ZSql = ZSql + "'" + Responsable25.Text + "',"
            ZSql = ZSql + "'" + Responsable26.Text + "',"
            ZSql = ZSql + "'" + Fecha21.Text + "',"
            ZSql = ZSql + "'" + Fecha22.Text + "',"
            ZSql = ZSql + "'" + Fecha23.Text + "',"
            ZSql = ZSql + "'" + Fecha24.Text + "',"
            ZSql = ZSql + "'" + Fecha25.Text + "',"
            ZSql = ZSql + "'" + Fecha26.Text + "',"
            ZSql = ZSql + "'" + Str$(Estado1.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado2.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado3.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado4.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado5.ListIndex) + "',"
            ZSql = ZSql + "'" + Str$(Estado6.ListIndex) + "',"
            ZSql = ZSql + "'" + Comentario211.Text + "',"
            ZSql = ZSql + "'" + Comentario212.Text + "',"
            ZSql = ZSql + "'" + Comentario221.Text + "',"
            ZSql = ZSql + "'" + Comentario222.Text + "',"
            ZSql = ZSql + "'" + Comentario231.Text + "',"
            ZSql = ZSql + "'" + Comentario232.Text + "',"
            ZSql = ZSql + "'" + Comentario241.Text + "',"
            ZSql = ZSql + "'" + Comentario242.Text + "',"
            ZSql = ZSql + "'" + Comentario251.Text + "',"
            ZSql = ZSql + "'" + Comentario252.Text + "',"
            ZSql = ZSql + "'" + Comentario261.Text + "',"
            ZSql = ZSql + "'" + Comentario262.Text + "',"
            ZSql = ZSql + "'" + Email1.Text + "',"
            ZSql = ZSql + "'" + Email2.Text + "',"
            ZSql = ZSql + "'" + Email3.Text + "',"
            ZSql = ZSql + "'" + Email4.Text + "',"
            ZSql = ZSql + "'" + Email5.Text + "',"
            ZSql = ZSql + "'" + Email6.Text + "',"
            ZSql = ZSql + "'" + Email7.Text + "',"
            ZSql = ZSql + "'" + Email8.Text + "',"
            ZSql = ZSql + "'" + Email9.Text + "',"
            ZSql = ZSql + "'" + Email10.Text + "',"
            ZSql = ZSql + "'" + DesResponsable1.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable2.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable3.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable4.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable5.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable6.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable11.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable12.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable13.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable14.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable15.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable16.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable21.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable22.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable23.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable24.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable25.Caption + "',"
            ZSql = ZSql + "'" + DesResponsable26.Caption + "',"
            ZSql = ZSql + "'" + Seccion.Text + "',"
            ZSql = ZSql + "'" + Estado.Text + "',"
            ZSql = ZSql + "'" + Estado1.Text + "',"
            ZSql = ZSql + "'" + Estado2.Text + "',"
            ZSql = ZSql + "'" + Estado3.Text + "',"
            ZSql = ZSql + "'" + Estado4.Text + "',"
            ZSql = ZSql + "'" + Estado5.Text + "',"
            ZSql = ZSql + "'" + Estado6.Text + "',"
            ZSql = ZSql + "'" + DesCentro.Caption + "')"
            
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
        End If
        
        Call CmdLimpiar_Click
        Centro.SetFocus
        
    End If
End Sub

Private Sub cmdDelete_Click()

    If Centro.Text <> "" And Numero.Text <> "" Then
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " CargaSac"
        ZSql = ZSql + " Where Centro = " + "'" + Centro.Text + "'"
        ZSql = ZSql + " and Numero = " + "'" + Numero.Text + "'"
        spCargaSac = ZSql
        Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
        If rstCargaSac.RecordCount > 0 Then
        
            rstCargaSac.Close
            
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
            
                
                Call CmdLimpiar_Click
                
            End If
        End If
        
    End If
    
    Orden.SetFocus
    
End Sub

Private Sub CmdLimpiar_Click()
    
    Centro.Text = ""
    DesCentro.Caption = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Origen.Text = ""
    Motivo.Text = ""
    
    Seccion.ListIndex = 0
    Estado.ListIndex = 0
    
    Descripcion1.Text = ""
    Descripcion2.Text = ""
    Descripcion3.Text = ""
    Descripcion4.Text = ""
    Descripcion5.Text = ""
    Descripcion6.Text = ""
    Descripcion7.Text = ""
    Descripcion8.Text = ""
    Descripcion9.Text = ""
    Descripcion10.Text = ""
    
    Causa1.Text = ""
    Causa2.Text = ""
    Causa3.Text = ""
    Causa4.Text = ""
    Causa5.Text = ""
    Causa6.Text = ""
    Causa7.Text = ""
    Causa8.Text = ""
    Causa9.Text = ""
    Causa10.Text = ""
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    Accion211.Text = ""
    Accion212.Text = ""
    Accion221.Text = ""
    Accion222.Text = ""
    Accion231.Text = ""
    Accion232.Text = ""
    Accion241.Text = ""
    Accion242.Text = ""
    Accion251.Text = ""
    Accion252.Text = ""
    Accion251.Text = ""
    Accion262.Text = ""
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    Email1.Text = ""
    Email2.Text = ""
    Email3.Text = ""
    Email4.Text = ""
    Email5.Text = ""
    Email6.Text = ""
    Email7.Text = ""
    Email8.Text = ""
    Email9.Text = ""
    Email10.Text = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Tablas.Tab = 0
    
    Centro.SetFocus
    
End Sub

Private Sub cmdClose_Click()

    PrgCargaSac.Hide
    Unload Me
    Menu.Show
    
End Sub



Sub Form_Load()

    Centro.Text = ""
    DesCentro.Caption = ""
    Numero.Text = ""
    Fecha.Text = "  /  /    "
    Origen.Text = ""
    Motivo.Text = ""
    
    Estado.Clear
    
    Estado.AddItem ""
    Estado.AddItem "IMPLEMENTACION"
    Estado.AddItem "VERIFICACION"
    Estado.AddItem "CERRADO"
    
    Estado.ListIndex = 0
    
    Seccion.Clear
    
    Seccion.AddItem ""
    Seccion.AddItem "AC"
    Seccion.AddItem "VENTAS"
    
    Seccion.ListIndex = 0
    
    Estado1.Clear
    
    Estado1.AddItem ""
    Estado1.AddItem "Pendiente"
    Estado1.AddItem "Efectiva"
    Estado1.AddItem "Cerrada"
    
    Estado2.Clear
    
    Estado2.AddItem ""
    Estado1.AddItem "Pendiente"
    Estado2.AddItem "Efectiva"
    Estado2.AddItem "Cerrada"
    
    Estado3.Clear
    
    Estado3.AddItem ""
    Estado3.AddItem "Pendiente"
    Estado3.AddItem "Efectiva"
    Estado3.AddItem "Cerrada"
    
    Estado4.Clear
    
    Estado4.AddItem ""
    Estado4.AddItem "Pendiente"
    Estado4.AddItem "Efectiva"
    Estado4.AddItem "Cerrada"
    
    Estado5.Clear
    
    Estado5.AddItem ""
    Estado5.AddItem "Pendiente"
    Estado5.AddItem "Efectiva"
    Estado5.AddItem "Cerrada"
    
    Estado6.Clear
    
    Estado6.AddItem ""
    Estado6.AddItem "Pendiente"
    Estado6.AddItem "Efectiva"
    Estado6.AddItem "Cerrada"
    
    Descripcion1.Text = ""
    Descripcion2.Text = ""
    Descripcion3.Text = ""
    Descripcion4.Text = ""
    Descripcion5.Text = ""
    Descripcion6.Text = ""
    Descripcion7.Text = ""
    Descripcion8.Text = ""
    Descripcion9.Text = ""
    Descripcion10.Text = ""
    
    Causa1.Text = ""
    Causa2.Text = ""
    Causa3.Text = ""
    Causa4.Text = ""
    Causa5.Text = ""
    Causa6.Text = ""
    Causa7.Text = ""
    Causa8.Text = ""
    Causa9.Text = ""
    Causa10.Text = ""
    
    Accion11.Text = ""
    Accion12.Text = ""
    Accion21.Text = ""
    Accion22.Text = ""
    Accion31.Text = ""
    Accion32.Text = ""
    Accion41.Text = ""
    Accion42.Text = ""
    Accion51.Text = ""
    Accion52.Text = ""
    Accion61.Text = ""
    Accion62.Text = ""
    
    Responsable1.Text = ""
    Responsable2.Text = ""
    Responsable3.Text = ""
    Responsable4.Text = ""
    Responsable5.Text = ""
    Responsable6.Text = ""
    
    DesResponsable1.Caption = ""
    DesResponsable2.Caption = ""
    DesResponsable3.Caption = ""
    DesResponsable4.Caption = ""
    DesResponsable5.Caption = ""
    DesResponsable6.Caption = ""

    Plazo1.Text = "  /  /    "
    Plazo2.Text = "  /  /    "
    Plazo3.Text = "  /  /    "
    Plazo4.Text = "  /  /    "
    Plazo5.Text = "  /  /    "
    Plazo6.Text = "  /  /    "
    
    Accion111.Text = ""
    Accion112.Text = ""
    Accion121.Text = ""
    Accion122.Text = ""
    Accion131.Text = ""
    Accion132.Text = ""
    Accion141.Text = ""
    Accion142.Text = ""
    Accion151.Text = ""
    Accion152.Text = ""
    Accion151.Text = ""
    Accion162.Text = ""
    
    Responsable11.Text = ""
    Responsable12.Text = ""
    Responsable13.Text = ""
    Responsable14.Text = ""
    Responsable15.Text = ""
    Responsable16.Text = ""
    
    DesResponsable11.Caption = ""
    DesResponsable12.Caption = ""
    DesResponsable13.Caption = ""
    DesResponsable14.Caption = ""
    DesResponsable15.Caption = ""
    DesResponsable16.Caption = ""

    Fecha1.Text = "  /  /    "
    Fecha2.Text = "  /  /    "
    Fecha3.Text = "  /  /    "
    Fecha4.Text = "  /  /    "
    Fecha5.Text = "  /  /    "
    Fecha6.Text = "  /  /    "
    
    Comentario11.Text = ""
    Comentario12.Text = ""
    Comentario21.Text = ""
    Comentario22.Text = ""
    Comentario31.Text = ""
    Comentario32.Text = ""
    Comentario41.Text = ""
    Comentario42.Text = ""
    Comentario51.Text = ""
    Comentario52.Text = ""
    Comentario61.Text = ""
    Comentario62.Text = ""
    
    
    
    Accion211.Text = ""
    Accion212.Text = ""
    Accion221.Text = ""
    Accion222.Text = ""
    Accion231.Text = ""
    Accion232.Text = ""
    Accion241.Text = ""
    Accion242.Text = ""
    Accion251.Text = ""
    Accion252.Text = ""
    Accion251.Text = ""
    Accion262.Text = ""
    
    Responsable21.Text = ""
    Responsable22.Text = ""
    Responsable23.Text = ""
    Responsable24.Text = ""
    Responsable25.Text = ""
    Responsable26.Text = ""
    
    DesResponsable21.Caption = ""
    DesResponsable22.Caption = ""
    DesResponsable23.Caption = ""
    DesResponsable24.Caption = ""
    DesResponsable25.Caption = ""
    DesResponsable26.Caption = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Comentario211.Text = ""
    Comentario212.Text = ""
    Comentario221.Text = ""
    Comentario222.Text = ""
    Comentario231.Text = ""
    Comentario232.Text = ""
    Comentario241.Text = ""
    Comentario242.Text = ""
    Comentario251.Text = ""
    Comentario252.Text = ""
    Comentario261.Text = ""
    Comentario262.Text = ""
    
    Email1.Text = ""
    Email2.Text = ""
    Email3.Text = ""
    Email4.Text = ""
    Email5.Text = ""
    Email6.Text = ""
    Email7.Text = ""
    Email8.Text = ""
    Email9.Text = ""
    Email10.Text = ""
    
    Estado1.ListIndex = 0
    Estado2.ListIndex = 0
    Estado3.ListIndex = 0
    Estado4.ListIndex = 0
    Estado5.ListIndex = 0
    Estado6.ListIndex = 0
    
    Rem Sql1 = "Select Max(Codigo) as [CodigoMayor]"
    Rem Sql2 = " FROM CentroSac"
    Rem spCentroSac = Sql1 + Sql2
    Rem Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
    Rem If rstCentroSac.RecordCount > 0 Then
    Rem     rstCentroSac.MoveLast
    Rem     ZCodigo = IIf(IsNull(rstCentroSac!CodigoMayor), "0", rstCentroSac!CodigoMayor)
    Rem     Codigo.Text = ZCodigo + 1
    Rem     rstCentroSac.Close
    Rem End If
    
    
End Sub

Private Sub Tablas_Click(PreviousTab As Integer)
    
    Select Case Tablas.Tab
        Case 0
            Descripcion1.SetFocus
        Case 1
            Accion11.SetFocus
        Case 2
            Accion111.Text = Accion11.Text
            Accion112.Text = Accion12.Text
            Accion121.Text = Accion21.Text
            Accion122.Text = Accion22.Text
            Accion131.Text = Accion31.Text
            Accion132.Text = Accion32.Text
            Accion141.Text = Accion41.Text
            Accion142.Text = Accion42.Text
            Accion151.Text = Accion51.Text
            Accion152.Text = Accion52.Text
            Accion161.Text = Accion61.Text
            Accion162.Text = Accion62.Text
            Responsable11.SetFocus
        Case 3
            Accion211.Text = Accion11.Text
            Accion212.Text = Accion12.Text
            Accion221.Text = Accion21.Text
            Accion222.Text = Accion22.Text
            Accion231.Text = Accion31.Text
            Accion232.Text = Accion32.Text
            Accion241.Text = Accion41.Text
            Accion242.Text = Accion42.Text
            Accion251.Text = Accion51.Text
            Accion252.Text = Accion52.Text
            Accion261.Text = Accion61.Text
            Accion262.Text = Accion62.Text
            Responsable21.SetFocus
        Case Else
    End Select
End Sub







Private Sub Centro_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Centro.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Where CentroSac.Codigo = " + "'" + Centro.Text + "'"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                DesCentro.Caption = Trim(rstCentroSac!Descripcion)
                rstCentroSac.Close
                Numero.SetFocus
            End If
        End If
    End If
    If KeyAscii = 27 Then
        Centro.Text = ""
        DesCentro.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Numero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Numero.Text <> "" Then
        
            ZSql = ""
            ZSql = ZSql + "Select Clave, Centro, Numero, Origen, Motivo, Seccion, Estado, Descripcion1, Descripcion2, Descripcion3, Descripcion4, Descripcion5, Descripcion6, Descripcion7, Descripcion8, Descripcion9, Descripcion10, Causa1, Causa2, Causa3, Causa4, Causa5, Causa6, Causa7, Causa8, Causa9, Causa10, Email1, Email2, Email3, Email4, Email5, Email6, Email7, Email8, Email9, Email10"
            ZSql = ZSql + " FROM CargaSac"
            ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
            ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
            spCargaSac = ZSql
            Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCargaSac.RecordCount > 0 Then
                
                Origen.Text = rstCargaSac!Origen
                Motivo.Text = rstCargaSac!Motivo
                Seccion.ListIndex = rstCargaSac!Seccion
                Estado.ListIndex = rstCargaSac!Estado
            
                Descripcion1.Text = rstCargaSac!Descripcion1
                Descripcion2.Text = rstCargaSac!Descripcion2
                Descripcion3.Text = rstCargaSac!Descripcion3
                Descripcion4.Text = rstCargaSac!Descripcion4
                Descripcion5.Text = rstCargaSac!Descripcion5
                Descripcion6.Text = rstCargaSac!Descripcion6
                Descripcion7.Text = rstCargaSac!Descripcion7
                Descripcion8.Text = rstCargaSac!Descripcion8
                Descripcion9.Text = rstCargaSac!Descripcion9
                Descripcion10.Text = rstCargaSac!Descripcion10
                
                Causa1.Text = rstCargaSac!Causa1
                Causa2.Text = rstCargaSac!Causa2
                Causa3.Text = rstCargaSac!Causa3
                Causa4.Text = rstCargaSac!Causa4
                Causa5.Text = rstCargaSac!Causa5
                Causa6.Text = rstCargaSac!Causa6
                Causa7.Text = rstCargaSac!Causa7
                Causa8.Text = rstCargaSac!Causa8
                Causa9.Text = rstCargaSac!Causa9
                Causa10.Text = rstCargaSac!Causa10
            
                Email1.Text = IIf(IsNull(rstCargaSac!Email1), "", rstCargaSac!Email1)
                Email2.Text = IIf(IsNull(rstCargaSac!Email2), "", rstCargaSac!Email2)
                Email3.Text = IIf(IsNull(rstCargaSac!Email3), "", rstCargaSac!Email3)
                Email4.Text = IIf(IsNull(rstCargaSac!Email4), "", rstCargaSac!Email4)
                Email5.Text = IIf(IsNull(rstCargaSac!Email5), "", rstCargaSac!Email5)
                Email6.Text = IIf(IsNull(rstCargaSac!Email6), "", rstCargaSac!Email6)
                Email7.Text = IIf(IsNull(rstCargaSac!Email7), "", rstCargaSac!Email7)
                Email8.Text = IIf(IsNull(rstCargaSac!Email8), "", rstCargaSac!Email8)
                Email9.Text = IIf(IsNull(rstCargaSac!Email9), "", rstCargaSac!Email9)
                Email10.Text = IIf(IsNull(rstCargaSac!Email10), "", rstCargaSac!Email10)
                
                rstCargaSac.Close
                
                ZSql = ""
                ZSql = ZSql + "Select Clave, Centro, Numero, Accion11, Accion12, Accion21, Accion22, Accion31, Accion32, Accion41, Accion42, Accion51, Accion52, Accion61, Accion62, Responsable1, Responsable2, Responsable3, Responsable4, Responsable5, Responsable6, Plazo1, Plazo2, Plazo3, Plazo4, Plazo5, Plazo6 "
                ZSql = ZSql + " FROM CargaSac"
                ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
                ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaSac.RecordCount > 0 Then
                    
                    Accion11.Text = rstCargaSac!Accion11
                    Accion12.Text = rstCargaSac!Accion12
                    Accion21.Text = rstCargaSac!Accion21
                    Accion22.Text = rstCargaSac!Accion22
                    Accion31.Text = rstCargaSac!Accion31
                    Accion32.Text = rstCargaSac!Accion32
                    Accion41.Text = rstCargaSac!Accion41
                    Accion42.Text = rstCargaSac!Accion42
                    Accion51.Text = rstCargaSac!Accion51
                    Accion52.Text = rstCargaSac!Accion52
                    Accion61.Text = rstCargaSac!Accion61
                    Accion62.Text = rstCargaSac!Accion62
                    
                    Responsable1.Text = rstCargaSac!Responsable1
                    Responsable2.Text = rstCargaSac!Responsable2
                    Responsable3.Text = rstCargaSac!Responsable3
                    Responsable4.Text = rstCargaSac!Responsable4
                    Responsable5.Text = rstCargaSac!Responsable5
                    Responsable6.Text = rstCargaSac!Responsable6
                    
                    Plazo1.Text = rstCargaSac!Plazo1
                    Plazo2.Text = rstCargaSac!Plazo2
                    Plazo3.Text = rstCargaSac!Plazo3
                    Plazo4.Text = rstCargaSac!Plazo4
                    Plazo5.Text = rstCargaSac!Plazo5
                    Plazo6.Text = rstCargaSac!Plazo6
            
                    rstCargaSac.Close
                End If
                
                
                ZSql = ""
                ZSql = ZSql + "Select Clave, Centro, Numero, Responsable11, Responsable12, Responsable13, Responsable14, Responsable15, Responsable16, Fecha1, Fecha2, Fecha3, Fecha4, Fecha5, Fecha6, Comentario11, Comentario12, Comentario21, Comentario22, Comentario31, Comentario32, Comentario41, Comentario42, Comentario51, Comentario52, Comentario61, Comentario62 "
                ZSql = ZSql + " FROM CargaSac"
                ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
                ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaSac.RecordCount > 0 Then
                    
                    Responsable11.Text = rstCargaSac!Responsable11
                    Responsable12.Text = rstCargaSac!Responsable12
                    Responsable13.Text = rstCargaSac!Responsable13
                    Responsable14.Text = rstCargaSac!Responsable14
                    Responsable15.Text = rstCargaSac!Responsable15
                    Responsable16.Text = rstCargaSac!Responsable16
                    
                    Fecha1.Text = rstCargaSac!Fecha1
                    Fecha2.Text = rstCargaSac!Fecha2
                    Fecha3.Text = rstCargaSac!Fecha3
                    Fecha4.Text = rstCargaSac!Fecha4
                    Fecha5.Text = rstCargaSac!Fecha5
                    Fecha6.Text = rstCargaSac!Fecha6
                    
                    Comentario11.Text = rstCargaSac!Comentario11
                    Comentario12.Text = rstCargaSac!Comentario12
                    Comentario21.Text = rstCargaSac!Comentario21
                    Comentario22.Text = rstCargaSac!Comentario22
                    Comentario31.Text = rstCargaSac!Comentario31
                    Comentario32.Text = rstCargaSac!Comentario32
                    Comentario41.Text = rstCargaSac!Comentario41
                    Comentario42.Text = rstCargaSac!Comentario42
                    Comentario51.Text = rstCargaSac!Comentario51
                    Comentario52.Text = rstCargaSac!Comentario52
                    Comentario61.Text = rstCargaSac!Comentario61
                    Comentario62.Text = rstCargaSac!Comentario62
            
                    rstCargaSac.Close
                End If
                
                
                
                ZSql = ""
                ZSql = ZSql + "Select Clave, Centro, Numero, Responsable21, Responsable22, Responsable23, Responsable24, Responsable25, Responsable26, Fecha21, Fecha22, Fecha23, Fecha24, Fecha25, Fecha26, Comentario211, Comentario212, Comentario221, Comentario222, Comentario231, Comentario232, Comentario241, Comentario242, Comentario251, Comentario252, Comentario261, Comentario262, Estado1, Estado2, Estado3, Estado4, Estado5, Estado6 "
                ZSql = ZSql + " FROM CargaSac"
                ZSql = ZSql + " Where CargaSac.Centro = " + "'" + Centro.Text + "'"
                ZSql = ZSql + " and CargaSac.Numero = " + "'" + Numero.Text + "'"
                spCargaSac = ZSql
                Set rstCargaSac = db.OpenRecordset(spCargaSac, dbOpenSnapshot, dbSQLPassThrough)
                If rstCargaSac.RecordCount > 0 Then
                    
                    Responsable21.Text = rstCargaSac!Responsable21
                    Responsable22.Text = rstCargaSac!Responsable22
                    Responsable23.Text = rstCargaSac!Responsable23
                    Responsable24.Text = rstCargaSac!Responsable24
                    Responsable25.Text = rstCargaSac!Responsable25
                    Responsable26.Text = rstCargaSac!Responsable26
                    
                    Fecha21.Text = rstCargaSac!Fecha21
                    Fecha22.Text = rstCargaSac!Fecha22
                    Fecha23.Text = rstCargaSac!Fecha23
                    Fecha24.Text = rstCargaSac!Fecha24
                    Fecha25.Text = rstCargaSac!Fecha25
                    Fecha26.Text = rstCargaSac!Fecha26
                    
                    Comentario211.Text = rstCargaSac!Comentario211
                    Comentario212.Text = rstCargaSac!Comentario212
                    Comentario221.Text = rstCargaSac!Comentario221
                    Comentario222.Text = rstCargaSac!Comentario222
                    Comentario231.Text = rstCargaSac!Comentario231
                    Comentario232.Text = rstCargaSac!Comentario232
                    Comentario241.Text = rstCargaSac!Comentario241
                    Comentario242.Text = rstCargaSac!Comentario242
                    Comentario251.Text = rstCargaSac!Comentario251
                    Comentario252.Text = rstCargaSac!Comentario252
                    Comentario261.Text = rstCargaSac!Comentario261
                    Comentario262.Text = rstCargaSac!Comentario262
                    
                    Estado1.ListIndex = rstCargaSac!Estado1
                    Estado2.ListIndex = rstCargaSac!Estado2
                    Estado3.ListIndex = rstCargaSac!Estado3
                    Estado4.ListIndex = rstCargaSac!Estado4
                    Estado5.ListIndex = rstCargaSac!Estado5
                    Estado6.ListIndex = rstCargaSac!Estado6
            
                    rstCargaSac.Close
                End If
                
                Call Imprime_Descripcion
                
                    Else
                    
                WCentro = Centro.Text
                WNumero = Numero.Text
                CmdLimpiar_Click
                Centro.Text = WCentro
                Numero.Text = WNumero
                Fecha.SetFocus
            End If
            
        End If
    End If
    If KeyAscii = 27 Then
        Numero.Text = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha.Text, Auxi)
        If Auxi = "S" Then
            Seccion.SetFocus
                Else
            Fecha.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha.Text = "  /  /    "
    End If
End Sub

Private Sub Seccion_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Estado.SetFocus
    End If
End Sub

Private Sub Estado_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Origen.SetFocus
    End If
End Sub

Private Sub Origen_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Motivo.SetFocus
    End If
    If KeyAscii = 27 Then
        Origen.Text = ""
    End If
End Sub

Private Sub Motivo_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    End If
    If KeyAscii = 27 Then
        Motivo.Text = ""
    End If
End Sub


Rem tab 1

Private Sub Descripcion1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa1.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion1.Text = ""
    End If
End Sub

Private Sub Causa1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion2.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa1.Text = ""
    End If
End Sub

Private Sub Descripcion2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa2.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion2.Text = ""
    End If
End Sub

Private Sub Causa2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion3.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa2.Text = ""
    End If
End Sub

Private Sub Descripcion3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa3.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion3.Text = ""
    End If
End Sub

Private Sub Causa3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion4.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa3.Text = ""
    End If
End Sub

Private Sub Descripcion4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa4.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion4.Text = ""
    End If
End Sub

Private Sub Causa4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion5.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa5.Text = ""
    End If
End Sub

Private Sub Descripcion5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa5.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion5.Text = ""
    End If
End Sub

Private Sub Causa5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion6.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa5.Text = ""
    End If
End Sub

Private Sub Descripcion6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa6.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion6.Text = ""
    End If
End Sub

Private Sub Causa6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion7.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa6.Text = ""
    End If
End Sub

Private Sub Descripcion7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa7.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion7.Text = ""
    End If
End Sub

Private Sub Causa7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion8.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa7.Text = ""
    End If
End Sub

Private Sub Descripcion8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa8.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion8.Text = ""
    End If
End Sub

Private Sub Causa8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion9.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa8.Text = ""
    End If
End Sub

Private Sub Descripcion9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa9.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion9.Text = ""
    End If
End Sub

Private Sub Causa9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion10.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa9.Text = ""
    End If
End Sub

Private Sub Descripcion10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Causa10.SetFocus
    End If
    If KeyAscii = 27 Then
        Descripcion10.Text = ""
    End If
End Sub

Private Sub Causa10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Descripcion1.SetFocus
    End If
    If KeyAscii = 27 Then
        Causa10.Text = ""
    End If
End Sub



Rem tab 2







Private Sub Accion11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion12.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion11.Text = ""
    End If
End Sub

Private Sub Accion12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable1.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion12.Text = ""
    End If
End Sub

Private Sub Responsable1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable1.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable1.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable1.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo1.SetFocus
            End If
                Else
            DesResponsable1.Caption = ""
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable1.Text = ""
        DesResponsable1.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo1.Text, Auxi)
        If Auxi = "S" Or Plazo1.Text = "  /  /    " Then
            Accion21.SetFocus
                Else
            Plazo1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo1.Text = "  /  /    "
    End If
End Sub








Private Sub Accion21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion22.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion21.Text = ""
    End If
End Sub

Private Sub Accion22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable2.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion22.Text = ""
    End If
End Sub

Private Sub Responsable2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable2.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable2.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable2.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo2.SetFocus
            End If
                Else
            DesResponsable2.Caption = ""
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable2.Text = ""
        DesResponsable2.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo2.Text, Auxi)
        If Auxi = "S" Or Plazo2.Text = "  /  /    " Then
            Accion31.SetFocus
                Else
            Plazo2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo2.Text = "  /  /    "
    End If
End Sub







Private Sub Accion31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion32.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion31.Text = ""
    End If
End Sub

Private Sub Accion32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable3.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion32.Text = ""
    End If
End Sub

Private Sub Responsable3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable3.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable3.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable3.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo3.SetFocus
            End If
                Else
            DesResponsable3.Caption = ""
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable3.Text = ""
        DesResponsable3.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo3.Text, Auxi)
        If Auxi = "S" Or Plazo3.Text = "  /  /    " Then
            Accion41.SetFocus
                Else
            Plazo3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo3.Text = "  /  /    "
    End If
End Sub







Private Sub Accion41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion42.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion41.Text = ""
    End If
End Sub

Private Sub Accion42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable4.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion42.Text = ""
    End If
End Sub

Private Sub Responsable4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable4.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable4.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable4.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo4.SetFocus
            End If
                Else
            DesResponsable4.Caption = ""
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable4.Text = ""
        DesResponsable4.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo4.Text, Auxi)
        If Auxi = "S" Or Plazo4.Text = "  /  /    " Then
            Accion51.SetFocus
                Else
            Plazo4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo4.Text = "  /  /    "
    End If
End Sub







Private Sub Accion51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion52.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion51.Text = ""
    End If
End Sub

Private Sub Accion52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable5.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion52.Text = ""
    End If
End Sub

Private Sub Responsable5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable5.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable5.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable5.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo5.SetFocus
            End If
                Else
            DesResponsable5.Caption = ""
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable5.Text = ""
        DesResponsable5.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo5.Text, Auxi)
        If Auxi = "S" Or Plazo5.Text = "  /  /    " Then
            Accion61.SetFocus
                Else
            Plazo5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo5.Text = "  /  /    "
    End If
End Sub







Private Sub Accion61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Accion62.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion61.Text = ""
    End If
End Sub

Private Sub Accion62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable6.SetFocus
    End If
    If KeyAscii = 27 Then
        Accion62.Text = ""
    End If
End Sub

Private Sub Responsable6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable6.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable6.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable6.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Plazo6.SetFocus
            End If
                Else
            DesResponsable6.Caption = ""
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable6.Text = ""
        DesResponsable6.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Plazo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Plazo6.Text, Auxi)
        If Auxi = "S" Or Plazo6.Text = "  /  /    " Then
            Accion11.SetFocus
                Else
            Plazo6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Plazo6.Text = "  /  /    "
    End If
End Sub





Rem TAB 3

Rem DADA
Rem DADA
Rem DADA









Private Sub Responsable11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable11.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable11.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable11.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha1.SetFocus
            End If
                Else
            DesResponsable11.Caption = ""
            Fecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable11.Text = ""
        DesResponsable11.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha1.Text, Auxi)
        If Auxi = "S" Or Fecha1.Text = "  /  /    " Then
            Comentario11.SetFocus
                Else
            Fecha1.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha1.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario12.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario11.Text = ""
    End If
End Sub

Private Sub Comentario12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable12.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario12.Text = ""
    End If
End Sub










Private Sub Responsable12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable12.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable12.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable12.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha2.SetFocus
            End If
                Else
            DesResponsable12.Caption = ""
            Fecha2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable12.Text = ""
        DesResponsable12.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha2.Text, Auxi)
        If Auxi = "S" Or Fecha2.Text = "  /  /    " Then
            Comentario21.SetFocus
                Else
            Fecha2.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha2.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario22.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario21.Text = ""
    End If
End Sub

Private Sub Comentario22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable13.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario22.Text = ""
    End If
End Sub










Private Sub Responsable13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable13.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable13.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable13.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha3.SetFocus
            End If
                Else
            DesResponsable13.Caption = ""
            Fecha3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable13.Text = ""
        DesResponsable13.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha3.Text, Auxi)
        If Auxi = "S" Or Fecha3.Text = "  /  /    " Then
            Comentario31.SetFocus
                Else
            Fecha3.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha3.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario31_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario32.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario31.Text = ""
    End If
End Sub

Private Sub Comentario32_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable14.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario32.Text = ""
    End If
End Sub










Private Sub Responsable14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable14.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable14.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable14.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha4.SetFocus
            End If
                Else
            DesResponsable14.Caption = ""
            Fecha4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable14.Text = ""
        DesResponsable14.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha4.Text, Auxi)
        If Auxi = "S" Or Fecha4.Text = "  /  /    " Then
            Comentario41.SetFocus
                Else
            Fecha4.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha4.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario41_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario42.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario41.Text = ""
    End If
End Sub

Private Sub Comentario42_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable15.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario42.Text = ""
    End If
End Sub










Private Sub Responsable15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable15.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable15.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable15.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha5.SetFocus
            End If
                Else
            DesResponsable15.Caption = ""
            Fecha5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable15.Text = ""
        DesResponsable15.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha5.Text, Auxi)
        If Auxi = "S" Or Fecha5.Text = "  /  /    " Then
            Comentario51.SetFocus
                Else
            Fecha5.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha5.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario51_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario52.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario51.Text = ""
    End If
End Sub

Private Sub Comentario52_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable16.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario52.Text = ""
    End If
End Sub










Private Sub Responsable16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable16.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable16.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable16.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha6.SetFocus
            End If
                Else
            DesResponsable16.Caption = ""
            Fecha6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable16.Text = ""
        DesResponsable16.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha6.Text, Auxi)
        If Auxi = "S" Or Fecha6.Text = "  /  /    " Then
            Comentario61.SetFocus
                Else
            Fecha6.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha6.Text = "  /  /    "
    End If
End Sub

Private Sub Comentario61_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario62.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario61.Text = ""
    End If
End Sub

Private Sub Comentario62_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable11.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario62.Text = ""
    End If
End Sub













Rem TAB 4

Rem DADA
Rem DADA
Rem DADA









Private Sub Responsable21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable21.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable21.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable21.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha21.SetFocus
            End If
                Else
            DesResponsable21.Caption = ""
            Fecha21.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable21.Text = ""
        DesResponsable21.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha21.Text, Auxi)
        If Auxi = "S" Or Fecha21.Text = "  /  /    " Then
            Estado1.SetFocus
                Else
            Fecha21.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha21.Text = "  /  /    "
    End If
End Sub

Private Sub Estado1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario211.SetFocus
    End If
End Sub

Private Sub Comentario211_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario212.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario211.Text = ""
    End If
End Sub

Private Sub Comentario212_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable22.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario212.Text = ""
    End If
End Sub


















Private Sub Responsable22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable22.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable22.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable22.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha22.SetFocus
            End If
                Else
            DesResponsable22.Caption = ""
            Fecha22.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable22.Text = ""
        DesResponsable22.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha22.Text, Auxi)
        If Auxi = "S" Or Fecha22.Text = "  /  /    " Then
            Estado2.SetFocus
                Else
            Fecha22.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha22.Text = "  /  /    "
    End If
End Sub

Private Sub Estado2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario221.SetFocus
    End If
End Sub

Private Sub Comentario221_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario222.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario221.Text = ""
    End If
End Sub

Private Sub Comentario222_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable23.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario222.Text = ""
    End If
End Sub


















Private Sub Responsable23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable23.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable23.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable23.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha23.SetFocus
            End If
                Else
            DesResponsable23.Caption = ""
            Fecha23.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable23.Text = ""
        DesResponsable23.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha23.Text, Auxi)
        If Auxi = "S" Or Fecha23.Text = "  /  /    " Then
            Estado3.SetFocus
                Else
            Fecha23.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha23.Text = "  /  /    "
    End If
End Sub

Private Sub Estado3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario231.SetFocus
    End If
End Sub

Private Sub Comentario231_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario232.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario231.Text = ""
    End If
End Sub

Private Sub Comentario232_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable24.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario232.Text = ""
    End If
End Sub


















Private Sub Responsable24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable24.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable24.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable24.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha24.SetFocus
            End If
                Else
            DesResponsable24.Caption = ""
            Fecha24.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable24.Text = ""
        DesResponsable24.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha24.Text, Auxi)
        If Auxi = "S" Or Fecha24.Text = "  /  /    " Then
            Estado4.SetFocus
                Else
            Fecha24.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha24.Text = "  /  /    "
    End If
End Sub

Private Sub Estado4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario241.SetFocus
    End If
End Sub

Private Sub Comentario241_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario242.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario241.Text = ""
    End If
End Sub

Private Sub Comentario242_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable25.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario242.Text = ""
    End If
End Sub


















Private Sub Responsable25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable25.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable25.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable25.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha25.SetFocus
            End If
                Else
            DesResponsable25.Caption = ""
            Fecha25.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable25.Text = ""
        DesResponsable25.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha25.Text, Auxi)
        If Auxi = "S" Or Fecha25.Text = "  /  /    " Then
            Estado5.SetFocus
                Else
            Fecha25.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha25.Text = "  /  /    "
    End If
End Sub

Private Sub Estado5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario251.SetFocus
    End If
End Sub

Private Sub Comentario251_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario252.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario251.Text = ""
    End If
End Sub

Private Sub Comentario252_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable26.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario212.Text = ""
    End If
End Sub


















Private Sub Responsable26_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(Responsable26.Text) <> 0 Then
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Where ResponsableSac.Codigo = " + "'" + Responsable26.Text + "'"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                DesResponsable26.Caption = Trim(rstResponsableSac!Descripcion)
                rstResponsableSac.Close
                Fecha26.SetFocus
            End If
                Else
            DesResponsable26.Caption = ""
            Fecha26.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Responsable26.Text = ""
        DesResponsable26.Caption = ""
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Fecha26_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Valida_fecha(Fecha26.Text, Auxi)
        If Auxi = "S" Or Fecha26.Text = "  /  /    " Then
            Estado6.SetFocus
                Else
            Fecha26.SetFocus
        End If
    End If
    If KeyAscii = 27 Then
        Fecha21.Text = "  /  /    "
    End If
End Sub

Private Sub Estado6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario261.SetFocus
    End If
End Sub

Private Sub Comentario261_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Comentario262.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario261.Text = ""
    End If
End Sub

Private Sub Comentario262_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Responsable21.SetFocus
    End If
    If KeyAscii = 27 Then
        Comentario262.Text = ""
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Visible = False
    Pantalla.Visible = False

     Opcion.Clear

     Opcion.AddItem "Centro"
     Opcion.AddItem "Responsables"

     Opcion.Visible = True
End Sub

Private Sub Opcion_Click()

    Opcion.Visible = False
     
    Dim IngresaItem As String

    Pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    Ayuda.Text = ""
    Ayuda.Visible = True
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstCentroSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstCentroSac.Close
            End If
        
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                            Pantalla.AddItem IngresaItem
                            IngresaItem = rstResponsableSac!Codigo
                            WIndice.AddItem IngresaItem
                            .MoveNext
                                Else
                            Exit Do
                        End If
                    Loop
                End With
                rstResponsableSac.Close
            End If
        
        Case Else
    End Select
            
    Ayuda.SetFocus
    Pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    Pantalla.Visible = False
    Select Case XIndice
        Case 0
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Centro.Text = WIndice.List(Indice)
            Call Centro_Keypress(13)
            
        Case 1
            Ayuda.Visible = False
            Indice = Pantalla.ListIndex
            Select Case ZZLugar
                Case 1
                    Responsable1.Text = WIndice.List(Indice)
                    Call Responsable1_Keypress(13)
                Case 2
                    Responsable2.Text = WIndice.List(Indice)
                    Call Responsable2_Keypress(13)
                Case 3
                    Responsable3.Text = WIndice.List(Indice)
                    Call Responsable3_Keypress(13)
                Case 4
                    Responsable4.Text = WIndice.List(Indice)
                    Call Responsable4_Keypress(13)
                Case 5
                    Responsable5.Text = WIndice.List(Indice)
                    Call Responsable5_Keypress(13)
                Case 6
                    Responsable6.Text = WIndice.List(Indice)
                    Call Responsable6_Keypress(13)
                Case 7
                    Responsable11.Text = WIndice.List(Indice)
                    Call Responsable11_Keypress(13)
                Case 8
                    Responsable12.Text = WIndice.List(Indice)
                    Call Responsable12_Keypress(13)
                Case 9
                    Responsable13.Text = WIndice.List(Indice)
                    Call Responsable13_Keypress(13)
                Case 10
                    Responsable14.Text = WIndice.List(Indice)
                    Call Responsable14_Keypress(13)
                Case 11
                    Responsable15.Text = WIndice.List(Indice)
                    Call Responsable15_Keypress(13)
                Case 12
                    Responsable16.Text = WIndice.List(Indice)
                    Call Responsable16_Keypress(13)
                Case 13
                    Responsable21.Text = WIndice.List(Indice)
                    Call Responsable21_Keypress(13)
                Case 14
                    Responsable22.Text = WIndice.List(Indice)
                    Call Responsable22_Keypress(13)
                Case 15
                    Responsable23.Text = WIndice.List(Indice)
                    Call Responsable23_Keypress(13)
                Case 16
                    Responsable24.Text = WIndice.List(Indice)
                    Call Responsable24_Keypress(13)
                Case 17
                    Responsable25.Text = WIndice.List(Indice)
                    Call Responsable25_Keypress(13)
                Case 18
                    Responsable26.Text = WIndice.List(Indice)
                    Call Responsable26_Keypress(13)
                Case Else
            End Select
            
        Case Else
    End Select
    
End Sub

Private Sub aYUDA_Keypress(KeyAscii As Integer)

    On Error GoTo WError
    
    If KeyAscii = 13 Then

    LugarAyuda = 0
    WIndice.Clear
    
    WEspacios = Len(Ayuda.Text)
    
    XIndice = Opcion.ListIndex
    
    
    Select Case XIndice
        Case 0
            Sql1 = "Select *"
            Sql2 = " FROM CentroSac"
            Sql3 = " Order by CentroSac.Codigo"
            spCentroSac = Sql1 + Sql2 + Sql3
            Set rstCentroSac = db.OpenRecordset(spCentroSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstCentroSac.RecordCount > 0 Then
                With rstCentroSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstCentroSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstCentroSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstCentroSac!Codigo) + " " + rstCentroSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstCentroSac!Codigo
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
                rstCentroSac.Close
            End If
                
        Case 1
            Sql1 = "Select *"
            Sql2 = " FROM ResponsableSac"
            Sql3 = " Order by ResponsableSac.Codigo"
            spResponsableSac = Sql1 + Sql2 + Sql3
            Set rstResponsableSac = db.OpenRecordset(spResponsableSac, dbOpenSnapshot, dbSQLPassThrough)
            If rstResponsableSac.RecordCount > 0 Then
                With rstResponsableSac
                    .MoveFirst
                    Do
                        If .EOF = False Then
                            da = Len(rstResponsableSac!Descripcion) - WEspacios
                            For aa = 1 To da + 1
                                If Left$(Ayuda.Text, WEspacios) = Mid$(rstResponsableSac!Descripcion, aa, WEspacios) Then
                                    IngresaItem = Str$(rstResponsableSac!Codigo) + " " + rstResponsableSac!Descripcion
                                    Pantalla.AddItem IngresaItem
                                    IngresaItem = rstResponsableSac!Codigo
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
                rstResponsableSac.Close
            End If
                
        Case Else
    End Select
    
    End If
    
    Exit Sub
    
WError:
    Resume Next

End Sub


Private Sub Centro_DblClick()

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 0
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable1_DblClick()

    ZZLugar = 1

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable2_DblClick()

    ZZLugar = 2

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable3_DblClick()

    ZZLugar = 3

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable4_DblClick()

    ZZLugar = 4

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable5_DblClick()

    ZZLugar = 5

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable6_DblClick()

    ZZLugar = 6

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub









Private Sub Responsable11_DblClick()

    ZZLugar = 7

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable12_DblClick()

    ZZLugar = 8

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable13_DblClick()

    ZZLugar = 9

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable14_DblClick()

    ZZLugar = 10

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable15_DblClick()

    ZZLugar = 11

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable16_DblClick()

    ZZLugar = 12

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub










Private Sub Responsable21_DblClick()

    ZZLugar = 13

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable22_DblClick()

    ZZLugar = 14

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable23_DblClick()

    ZZLugar = 15

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable24_DblClick()

    ZZLugar = 16

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable25_DblClick()

    ZZLugar = 17

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub

Private Sub Responsable26_DblClick()

    ZZLugar = 18

    Opcion.Clear
    Opcion.AddItem "Centros"
    Opcion.AddItem "Responsable"
    Rem Opcion.Visible = True
    Opcion.ListIndex = 1
    
    Rem Call Opcion_Click

End Sub










