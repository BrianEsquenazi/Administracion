VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form PrgEspecifiUnifica 
   Caption         =   "Ingreso de Especificaciones de Materia Prima (Unificado)"
   ClientHeight    =   8295
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   11805
   LinkTopic       =   "Form2"
   ScaleHeight     =   8295
   ScaleWidth      =   11805
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9600
      TabIndex        =   246
      Top             =   7440
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Cas 
      Height          =   285
      Left            =   8880
      MaxLength       =   50
      TabIndex        =   180
      Text            =   " "
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox DescripcionIngles 
      Height          =   285
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   177
      Text            =   " "
      Top             =   720
      Width           =   5895
   End
   Begin VB.CommandButton ImprimeII 
      Caption         =   "Especif. Ingles"
      Height          =   495
      Left            =   10800
      TabIndex        =   176
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton GrabaII 
      Caption         =   "Graba Idioma"
      Height          =   495
      Left            =   10800
      TabIndex        =   175
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Idioma 
      Caption         =   "Cambio Idioma"
      Height          =   495
      Left            =   10680
      TabIndex        =   174
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox ControlCambio 
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
      Left            =   2040
      MaxLength       =   100
      TabIndex        =   152
      Text            =   " "
      Top             =   5760
      Width           =   5760
   End
   Begin VB.Frame XClave 
      Height          =   1935
      Left            =   2880
      TabIndex        =   37
      Top             =   5880
      Visible         =   0   'False
      Width           =   3735
      Begin VB.TextBox WClave 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   39
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton CancelaGraba 
         Caption         =   "Cancela Ingreso"
         Height          =   255
         Left            =   720
         TabIndex        =   38
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ingrese su Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         TabIndex        =   40
         Top             =   240
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4575
      Left            =   120
      TabIndex        =   41
      Top             =   1080
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   8070
      _Version        =   327680
      TabHeight       =   520
      TabCaption(0)   =   "Especificacion 1 - 10"
      TabPicture(0)   =   "especifiunifica.frx":0000
      Tab(0).ControlCount=   65
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Descri10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Descri9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Descri8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Descri7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Descri6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Descri5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Descri4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Descri3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "descri2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Descri1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblDescri"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblensayo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Titulo"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Ensayo10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Ensayo9"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Ensayo8"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Ensayo7"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Ensayo6"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Ensayo5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Ensayo4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Ensayo3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Ensayo2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Ensayo1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "valor10"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "valor9"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "valor8"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "valor7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "valor6"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "valor5"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "valor4"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Valor3"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "valor2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Valor1"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Desde1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Hasta1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Desde2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Hasta2"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Desde3"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Hasta3"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Desde4"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Hasta4"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Desde5"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Hasta5"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Desde6"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Hasta6"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Desde7"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Hasta7"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "Desde8"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "Hasta8"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "Desde9"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "Hasta9"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "Desde10"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "Hasta10"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "IValor1"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "IValor2"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "IValor3"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "IValor4"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "IValor5"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "IValor6"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "IValor7"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "IValor8"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "IValor9"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "IValor10"
      Tab(0).Control(64).Enabled=   0   'False
      TabCaption(1)   =   "Especificacion 11  - 20"
      TabPicture(1)   =   "especifiunifica.frx":001C
      Tab(1).ControlCount=   65
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "IValor20"
      Tab(1).Control(0).Enabled=   -1  'True
      Tab(1).Control(1)=   "IValor19"
      Tab(1).Control(1).Enabled=   -1  'True
      Tab(1).Control(2)=   "IValor18"
      Tab(1).Control(2).Enabled=   -1  'True
      Tab(1).Control(3)=   "IValor17"
      Tab(1).Control(3).Enabled=   -1  'True
      Tab(1).Control(4)=   "IValor16"
      Tab(1).Control(4).Enabled=   -1  'True
      Tab(1).Control(5)=   "IValor15"
      Tab(1).Control(5).Enabled=   -1  'True
      Tab(1).Control(6)=   "IValor14"
      Tab(1).Control(6).Enabled=   -1  'True
      Tab(1).Control(7)=   "IValor13"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "IValor12"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "IValor11"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Hasta20"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Desde20"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Hasta19"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Desde19"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Hasta18"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Desde18"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Hasta17"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Desde17"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Hasta16"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Desde16"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Hasta15"
      Tab(1).Control(20).Enabled=   -1  'True
      Tab(1).Control(21)=   "Desde15"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "Hasta14"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Desde14"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Hasta13"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Desde13"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Hasta12"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Desde12"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Hasta11"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Desde11"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Ensayo20"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Ensayo19"
      Tab(1).Control(31).Enabled=   -1  'True
      Tab(1).Control(32)=   "Ensayo18"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Ensayo17"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Ensayo16"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Ensayo15"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Ensayo14"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Ensayo13"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Ensayo12"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Ensayo11"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Valor20"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Valor19"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Valor18"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Valor17"
      Tab(1).Control(43).Enabled=   -1  'True
      Tab(1).Control(44)=   "Valor16"
      Tab(1).Control(44).Enabled=   -1  'True
      Tab(1).Control(45)=   "Valor15"
      Tab(1).Control(45).Enabled=   -1  'True
      Tab(1).Control(46)=   "Valor14"
      Tab(1).Control(46).Enabled=   -1  'True
      Tab(1).Control(47)=   "Valor13"
      Tab(1).Control(47).Enabled=   -1  'True
      Tab(1).Control(48)=   "Valor12"
      Tab(1).Control(48).Enabled=   -1  'True
      Tab(1).Control(49)=   "Valor11"
      Tab(1).Control(49).Enabled=   -1  'True
      Tab(1).Control(50)=   "Label14"
      Tab(1).Control(50).Enabled=   0   'False
      Tab(1).Control(51)=   "Label12"
      Tab(1).Control(51).Enabled=   0   'False
      Tab(1).Control(52)=   "Descri20"
      Tab(1).Control(52).Enabled=   0   'False
      Tab(1).Control(53)=   "Descri19"
      Tab(1).Control(53).Enabled=   0   'False
      Tab(1).Control(54)=   "Descri18"
      Tab(1).Control(54).Enabled=   0   'False
      Tab(1).Control(55)=   "Descri17"
      Tab(1).Control(55).Enabled=   0   'False
      Tab(1).Control(56)=   "Descri16"
      Tab(1).Control(56).Enabled=   0   'False
      Tab(1).Control(57)=   "Descri15"
      Tab(1).Control(57).Enabled=   0   'False
      Tab(1).Control(58)=   "Descri14"
      Tab(1).Control(58).Enabled=   0   'False
      Tab(1).Control(59)=   "Descri13"
      Tab(1).Control(59).Enabled=   0   'False
      Tab(1).Control(60)=   "Descri12"
      Tab(1).Control(60).Enabled=   0   'False
      Tab(1).Control(61)=   "Descri11"
      Tab(1).Control(61).Enabled=   0   'False
      Tab(1).Control(62)=   "Label8"
      Tab(1).Control(62).Enabled=   0   'False
      Tab(1).Control(63)=   "Label7"
      Tab(1).Control(63).Enabled=   0   'False
      Tab(1).Control(64)=   "TituloII"
      Tab(1).Control(64).Enabled=   0   'False
      TabCaption(2)   =   "Especificacion 21  - 30"
      TabPicture(2)   =   "especifiunifica.frx":0038
      Tab(2).ControlCount=   65
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label15"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Descri30"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Descri29"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Descri28"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Descri27"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Descri26"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Descri25"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Descri24"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Descri23"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Descri22"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Descri21"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Label27"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Label28"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "TituloIII"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Valor21"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "Valor22"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Valor23"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Valor24"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "Valor25"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "Valor26"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Valor27"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "Valor28"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Valor29"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Valor30"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "IValor30"
      Tab(2).Control(25).Enabled=   -1  'True
      Tab(2).Control(26)=   "IValor29"
      Tab(2).Control(26).Enabled=   -1  'True
      Tab(2).Control(27)=   "IValor28"
      Tab(2).Control(27).Enabled=   -1  'True
      Tab(2).Control(28)=   "IValor27"
      Tab(2).Control(28).Enabled=   -1  'True
      Tab(2).Control(29)=   "IValor26"
      Tab(2).Control(29).Enabled=   -1  'True
      Tab(2).Control(30)=   "IValor25"
      Tab(2).Control(30).Enabled=   -1  'True
      Tab(2).Control(31)=   "IValor24"
      Tab(2).Control(31).Enabled=   -1  'True
      Tab(2).Control(32)=   "IValor23"
      Tab(2).Control(32).Enabled=   -1  'True
      Tab(2).Control(33)=   "IValor22"
      Tab(2).Control(33).Enabled=   -1  'True
      Tab(2).Control(34)=   "IValor21"
      Tab(2).Control(34).Enabled=   -1  'True
      Tab(2).Control(35)=   "Hasta30"
      Tab(2).Control(35).Enabled=   -1  'True
      Tab(2).Control(36)=   "Desde30"
      Tab(2).Control(36).Enabled=   -1  'True
      Tab(2).Control(37)=   "Hasta29"
      Tab(2).Control(37).Enabled=   -1  'True
      Tab(2).Control(38)=   "Desde29"
      Tab(2).Control(38).Enabled=   -1  'True
      Tab(2).Control(39)=   "Hasta28"
      Tab(2).Control(39).Enabled=   -1  'True
      Tab(2).Control(40)=   "Desde28"
      Tab(2).Control(40).Enabled=   -1  'True
      Tab(2).Control(41)=   "Hasta27"
      Tab(2).Control(41).Enabled=   -1  'True
      Tab(2).Control(42)=   "Desde27"
      Tab(2).Control(42).Enabled=   -1  'True
      Tab(2).Control(43)=   "Hasta26"
      Tab(2).Control(43).Enabled=   -1  'True
      Tab(2).Control(44)=   "Desde26"
      Tab(2).Control(44).Enabled=   -1  'True
      Tab(2).Control(45)=   "Hasta25"
      Tab(2).Control(45).Enabled=   -1  'True
      Tab(2).Control(46)=   "Desde25"
      Tab(2).Control(46).Enabled=   -1  'True
      Tab(2).Control(47)=   "Hasta24"
      Tab(2).Control(47).Enabled=   -1  'True
      Tab(2).Control(48)=   "Desde24"
      Tab(2).Control(48).Enabled=   -1  'True
      Tab(2).Control(49)=   "Hasta23"
      Tab(2).Control(49).Enabled=   -1  'True
      Tab(2).Control(50)=   "Desde23"
      Tab(2).Control(50).Enabled=   -1  'True
      Tab(2).Control(51)=   "Hasta22"
      Tab(2).Control(51).Enabled=   -1  'True
      Tab(2).Control(52)=   "Desde22"
      Tab(2).Control(52).Enabled=   -1  'True
      Tab(2).Control(53)=   "Hasta21"
      Tab(2).Control(53).Enabled=   -1  'True
      Tab(2).Control(54)=   "Desde21"
      Tab(2).Control(54).Enabled=   -1  'True
      Tab(2).Control(55)=   "Ensayo30"
      Tab(2).Control(55).Enabled=   -1  'True
      Tab(2).Control(56)=   "Ensayo29"
      Tab(2).Control(56).Enabled=   -1  'True
      Tab(2).Control(57)=   "Ensayo28"
      Tab(2).Control(57).Enabled=   -1  'True
      Tab(2).Control(58)=   "Ensayo27"
      Tab(2).Control(58).Enabled=   -1  'True
      Tab(2).Control(59)=   "Ensayo26"
      Tab(2).Control(59).Enabled=   -1  'True
      Tab(2).Control(60)=   "Ensayo25"
      Tab(2).Control(60).Enabled=   -1  'True
      Tab(2).Control(61)=   "Ensayo24"
      Tab(2).Control(61).Enabled=   -1  'True
      Tab(2).Control(62)=   "Ensayo23"
      Tab(2).Control(62).Enabled=   -1  'True
      Tab(2).Control(63)=   "Ensayo22"
      Tab(2).Control(63).Enabled=   -1  'True
      Tab(2).Control(64)=   "Ensayo21"
      Tab(2).Control(64).Enabled=   -1  'True
      Begin VB.TextBox Ensayo21 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   220
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Ensayo22 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   219
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo23 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   218
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo24 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   217
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo25 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   216
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo26 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   215
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo27 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   214
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo28 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   213
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo29 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   212
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo30 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   211
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Desde21 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   210
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Hasta21 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   209
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Desde22 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   208
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Hasta22 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   207
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Desde23 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   206
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Hasta23 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   205
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Desde24 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   204
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Hasta24 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   203
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Desde25 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   202
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Hasta25 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   201
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Desde26 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   200
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Hasta26 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   199
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Desde27 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   198
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Hasta27 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   197
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Desde28 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   196
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Hasta28 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   195
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Desde29 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   194
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Hasta29 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   193
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Desde30 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   192
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox Hasta30 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   191
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox IValor21 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   190
         Text            =   " "
         Top             =   840
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor22 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   189
         Text            =   " "
         Top             =   1200
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor23 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   188
         Text            =   " "
         Top             =   1560
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor24 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   187
         Text            =   " "
         Top             =   1920
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor25 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   186
         Text            =   " "
         Top             =   2280
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor26 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   185
         Text            =   " "
         Top             =   2640
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor27 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   184
         Text            =   " "
         Top             =   3000
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor28 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   183
         Text            =   " "
         Top             =   3360
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor29 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   182
         Text            =   " "
         Top             =   3720
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor30 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   181
         Text            =   " "
         Top             =   4080
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor20 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   173
         Text            =   " "
         Top             =   4080
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor19 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   172
         Text            =   " "
         Top             =   3720
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor18 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   171
         Text            =   " "
         Top             =   3360
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor17 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   170
         Text            =   " "
         Top             =   3000
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor16 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   169
         Text            =   " "
         Top             =   2640
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor15 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   168
         Text            =   " "
         Top             =   2280
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor14 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   167
         Text            =   " "
         Top             =   1920
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor13 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   166
         Text            =   " "
         Top             =   1560
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor12 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   165
         Text            =   " "
         Top             =   1200
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor11 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   164
         Text            =   " "
         Top             =   840
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor10 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   163
         Text            =   " "
         Top             =   4080
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor9 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   162
         Text            =   " "
         Top             =   3720
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor8 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   161
         Text            =   " "
         Top             =   3360
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor7 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   160
         Text            =   " "
         Top             =   3000
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor6 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   159
         Text            =   " "
         Top             =   2640
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor5 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   158
         Text            =   " "
         Top             =   2280
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor4 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   157
         Text            =   " "
         Top             =   1920
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor3 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   156
         Text            =   " "
         Top             =   1560
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor2 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   155
         Text            =   " "
         Top             =   1200
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox IValor1 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   154
         Text            =   " "
         Top             =   840
         Visible         =   0   'False
         Width           =   5040
      End
      Begin VB.TextBox Hasta20 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   149
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox Desde20 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   148
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox Hasta19 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   147
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Desde19 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   146
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Hasta18 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   145
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Desde18 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   144
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Hasta17 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   143
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Desde17 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   142
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Hasta16 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   141
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Desde16 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   140
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Hasta15 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   139
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Desde15 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   138
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Hasta14 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   137
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Desde14 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   136
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Hasta13 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   135
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Desde13 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   134
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Hasta12 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   133
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Desde12 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   132
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Hasta11 
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
         Left            =   -64320
         MaxLength       =   8
         TabIndex        =   131
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Desde11 
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
         Left            =   -65160
         MaxLength       =   8
         TabIndex        =   130
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Hasta10 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   127
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox Desde10 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   126
         Text            =   " "
         Top             =   4080
         Width           =   840
      End
      Begin VB.TextBox Hasta9 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   125
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Desde9 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   124
         Text            =   " "
         Top             =   3720
         Width           =   840
      End
      Begin VB.TextBox Hasta8 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   123
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Desde8 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   122
         Text            =   " "
         Top             =   3360
         Width           =   840
      End
      Begin VB.TextBox Hasta7 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   121
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Desde7 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   120
         Text            =   " "
         Top             =   3000
         Width           =   840
      End
      Begin VB.TextBox Hasta6 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   119
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Desde6 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   118
         Text            =   " "
         Top             =   2640
         Width           =   840
      End
      Begin VB.TextBox Hasta5 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   117
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Desde5 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   116
         Text            =   " "
         Top             =   2280
         Width           =   840
      End
      Begin VB.TextBox Hasta4 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   115
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Desde4 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   114
         Text            =   " "
         Top             =   1920
         Width           =   840
      End
      Begin VB.TextBox Hasta3 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   113
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Desde3 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   112
         Text            =   " "
         Top             =   1560
         Width           =   840
      End
      Begin VB.TextBox Hasta2 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   111
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Desde2 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   110
         Text            =   " "
         Top             =   1200
         Width           =   840
      End
      Begin VB.TextBox Hasta1 
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
         Left            =   10680
         MaxLength       =   8
         TabIndex        =   109
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Desde1 
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
         Left            =   9840
         MaxLength       =   8
         TabIndex        =   108
         Text            =   " "
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Ensayo20 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   94
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Ensayo19 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   93
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo18 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   92
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo17 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   91
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo16 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   90
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo15 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   89
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo14 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   88
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo13 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   87
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   86
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo11 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   -74880
         MaxLength       =   4
         TabIndex        =   85
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Valor20 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   84
         Text            =   " "
         Top             =   4080
         Width           =   5040
      End
      Begin VB.TextBox Valor19 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   83
         Text            =   " "
         Top             =   3720
         Width           =   5040
      End
      Begin VB.TextBox Valor18 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   82
         Text            =   " "
         Top             =   3360
         Width           =   5040
      End
      Begin VB.TextBox Valor17 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   81
         Text            =   " "
         Top             =   3000
         Width           =   5040
      End
      Begin VB.TextBox Valor16 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   80
         Text            =   " "
         Top             =   2640
         Width           =   5040
      End
      Begin VB.TextBox Valor15 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   79
         Text            =   " "
         Top             =   2280
         Width           =   5040
      End
      Begin VB.TextBox Valor14 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   78
         Text            =   " "
         Top             =   1920
         Width           =   5040
      End
      Begin VB.TextBox Valor13 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   77
         Text            =   " "
         Top             =   1560
         Width           =   5040
      End
      Begin VB.TextBox Valor12 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   76
         Text            =   " "
         Top             =   1200
         Width           =   5040
      End
      Begin VB.TextBox Valor11 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   75
         Text            =   " "
         Top             =   840
         Width           =   5040
      End
      Begin VB.TextBox Valor1 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   73
         Text            =   " "
         Top             =   840
         Width           =   5040
      End
      Begin VB.TextBox valor2 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   72
         Text            =   " "
         Top             =   1200
         Width           =   5040
      End
      Begin VB.TextBox Valor3 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   71
         Text            =   " "
         Top             =   1560
         Width           =   5040
      End
      Begin VB.TextBox valor4 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   70
         Text            =   " "
         Top             =   1920
         Width           =   5040
      End
      Begin VB.TextBox valor5 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   69
         Text            =   " "
         Top             =   2280
         Width           =   5040
      End
      Begin VB.TextBox valor6 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   68
         Text            =   " "
         Top             =   2640
         Width           =   5040
      End
      Begin VB.TextBox valor7 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   67
         Text            =   " "
         Top             =   3000
         Width           =   5040
      End
      Begin VB.TextBox valor8 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   66
         Text            =   " "
         Top             =   3360
         Width           =   5040
      End
      Begin VB.TextBox valor9 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   65
         Text            =   " "
         Top             =   3720
         Width           =   5040
      End
      Begin VB.TextBox valor10 
         Height          =   285
         Left            =   4800
         MaxLength       =   70
         TabIndex        =   64
         Text            =   " "
         Top             =   4080
         Width           =   5040
      End
      Begin VB.TextBox Ensayo1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   51
         Text            =   " "
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox Ensayo2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   50
         Text            =   " "
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox Ensayo3 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   49
         Text            =   " "
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox Ensayo4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   48
         Text            =   " "
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox Ensayo5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   47
         Text            =   " "
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox Ensayo6 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         MaxLength       =   4
         TabIndex        =   46
         Text            =   " "
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox Ensayo7 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   45
         Text            =   " "
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox Ensayo8 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   44
         Text            =   " "
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox Ensayo9 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   43
         Text            =   " "
         Top             =   3720
         Width           =   735
      End
      Begin VB.TextBox Ensayo10 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         MaxLength       =   4
         TabIndex        =   42
         Text            =   " "
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Valor30 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   236
         Text            =   " "
         Top             =   4080
         Width           =   5040
      End
      Begin VB.TextBox Valor29 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   237
         Text            =   " "
         Top             =   3720
         Width           =   5040
      End
      Begin VB.TextBox Valor28 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   238
         Text            =   " "
         Top             =   3360
         Width           =   5040
      End
      Begin VB.TextBox Valor27 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   239
         Text            =   " "
         Top             =   3000
         Width           =   5040
      End
      Begin VB.TextBox Valor26 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   240
         Text            =   " "
         Top             =   2640
         Width           =   5040
      End
      Begin VB.TextBox Valor25 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   241
         Text            =   " "
         Top             =   2280
         Width           =   5040
      End
      Begin VB.TextBox Valor24 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   242
         Text            =   " "
         Top             =   1920
         Width           =   5040
      End
      Begin VB.TextBox Valor23 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   243
         Text            =   " "
         Top             =   1560
         Width           =   5040
      End
      Begin VB.TextBox Valor22 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   244
         Text            =   " "
         Top             =   1200
         Width           =   5040
      End
      Begin VB.TextBox Valor21 
         Height          =   285
         Left            =   -70200
         MaxLength       =   70
         TabIndex        =   245
         Text            =   " "
         Top             =   840
         Width           =   5040
      End
      Begin VB.Label TituloIII 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         TabIndex        =   235
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   234
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Left            =   -74040
         TabIndex        =   233
         Top             =   480
         Width           =   3840
      End
      Begin VB.Label Descri21 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   232
         Top             =   840
         Width           =   3840
      End
      Begin VB.Label Descri22 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   231
         Top             =   1200
         Width           =   3840
      End
      Begin VB.Label Descri23 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   230
         Top             =   1560
         Width           =   3840
      End
      Begin VB.Label Descri24 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   229
         Top             =   1920
         Width           =   3840
      End
      Begin VB.Label Descri25 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   228
         Top             =   2280
         Width           =   3840
      End
      Begin VB.Label Descri26 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   227
         Top             =   2640
         Width           =   3840
      End
      Begin VB.Label Descri27 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   226
         Top             =   3000
         Width           =   3840
      End
      Begin VB.Label Descri28 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   225
         Top             =   3360
         Width           =   3840
      End
      Begin VB.Label Descri29 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   224
         Top             =   3720
         Width           =   3840
      End
      Begin VB.Label Descri30 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   223
         Top             =   4080
         Width           =   3840
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
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
         Left            =   -65160
         TabIndex        =   222
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta"
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
         Left            =   -64320
         TabIndex        =   221
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta"
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
         Left            =   -64320
         TabIndex        =   151
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
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
         Left            =   -65160
         TabIndex        =   150
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Hasta"
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
         Left            =   10680
         TabIndex        =   129
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Desde"
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
         Left            =   9840
         TabIndex        =   128
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Descri20 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   107
         Top             =   4080
         Width           =   3840
      End
      Begin VB.Label Descri19 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   106
         Top             =   3720
         Width           =   3840
      End
      Begin VB.Label Descri18 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   105
         Top             =   3360
         Width           =   3840
      End
      Begin VB.Label Descri17 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   104
         Top             =   3000
         Width           =   3840
      End
      Begin VB.Label Descri16 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   103
         Top             =   2640
         Width           =   3840
      End
      Begin VB.Label Descri15 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   102
         Top             =   2280
         Width           =   3840
      End
      Begin VB.Label Descri14 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   101
         Top             =   1920
         Width           =   3840
      End
      Begin VB.Label Descri13 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   100
         Top             =   1560
         Width           =   3840
      End
      Begin VB.Label Descri12 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   99
         Top             =   1200
         Width           =   3840
      End
      Begin VB.Label Descri11 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   -74040
         TabIndex        =   98
         Top             =   840
         Width           =   3840
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Left            =   -74040
         TabIndex        =   97
         Top             =   480
         Width           =   3840
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   -74880
         TabIndex        =   96
         Top             =   480
         Width           =   735
      End
      Begin VB.Label TituloII 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         TabIndex        =   95
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label Titulo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Valor Standard"
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
         Left            =   4800
         TabIndex        =   74
         Top             =   480
         Width           =   5040
      End
      Begin VB.Label lblensayo 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ensayo"
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
         Left            =   120
         TabIndex        =   63
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblDescri 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripcion"
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
         Left            =   960
         TabIndex        =   62
         Top             =   480
         Width           =   3840
      End
      Begin VB.Label Descri1 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   61
         Top             =   840
         Width           =   3840
      End
      Begin VB.Label descri2 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   60
         Top             =   1200
         Width           =   3840
      End
      Begin VB.Label Descri3 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   59
         Top             =   1560
         Width           =   3840
      End
      Begin VB.Label Descri4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   58
         Top             =   1920
         Width           =   3840
      End
      Begin VB.Label Descri5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   57
         Top             =   2280
         Width           =   3840
      End
      Begin VB.Label Descri6 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   56
         Top             =   2640
         Width           =   3840
      End
      Begin VB.Label Descri7 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   55
         Top             =   3000
         Width           =   3840
      End
      Begin VB.Label Descri8 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   54
         Top             =   3360
         Width           =   3840
      End
      Begin VB.Label Descri9 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   53
         Top             =   3720
         Width           =   3840
      End
      Begin VB.Label Descri10 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Left            =   960
         TabIndex        =   52
         Top             =   4080
         Width           =   3840
      End
   End
   Begin VB.TextBox Fecha 
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
      Left            =   8880
      MaxLength       =   50
      TabIndex        =   34
      Text            =   " "
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Version 
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
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   33
      Text            =   " "
      Top             =   360
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Codigo 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   0
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
      Mask            =   "AA-###-###"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport lista 
      Left            =   6120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "wespec1Unifica.rpt"
      GroupSelectionFormula=   " "
      DiscardSavedData=   -1  'True
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   4320
      TabIndex        =   19
      Top             =   6000
      Visible         =   0   'False
      Width           =   3135
      Begin MSMask.MaskEdBox Hasta 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox Desde 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         MaxLength       =   10
         Mask            =   "AA-###-###"
         PromptChar      =   " "
      End
      Begin VB.OptionButton ImpreListado 
         Caption         =   "Option2"
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton ImprePantalla 
         Caption         =   "Option1"
         Height          =   195
         Left            =   1920
         TabIndex        =   24
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton Cancela 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   960
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton Acepta 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Impresora"
         Height          =   255
         Left            =   2160
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Pantalla"
         Height          =   255
         Left            =   2160
         TabIndex        =   26
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta  Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Desde Codigo"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.ListBox Opcion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   1080
      TabIndex        =   18
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox imprime 
      Height          =   285
      Left            =   4800
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox WIndice 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   6240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox pantalla 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "especifiunifica.frx":0054
      Left            =   120
      List            =   "especifiunifica.frx":005B
      TabIndex        =   13
      Top             =   6120
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.CommandButton Listado 
      Caption         =   "Listado"
      Height          =   255
      Left            =   7920
      TabIndex        =   12
      Top             =   7200
      Width           =   975
   End
   Begin VB.CommandButton Consulta 
      Caption         =   "Consulta"
      Height          =   255
      Left            =   7920
      TabIndex        =   11
      Top             =   6840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      Height          =   1335
      Left            =   9120
      TabIndex        =   6
      Top             =   5760
      Width           =   1575
      Begin VB.CommandButton Anterior 
         Caption         =   "Reg. Anterior"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Siguiente 
         Caption         =   "Reg. Siguiente"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton Ultimo 
         Caption         =   "Ultimo Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Primer 
         Caption         =   "Primer Reg."
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "Limpiar"
      Height          =   300
      Left            =   7920
      TabIndex        =   5
      Top             =   6480
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   7920
      TabIndex        =   4
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Eliminar"
      Enabled         =   0   'False
      Height          =   300
      Left            =   7920
      TabIndex        =   3
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Agregar"
      Height          =   300
      Left            =   7920
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cas"
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
      Index           =   5
      Left            =   8160
      TabIndex        =   179
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Desc. Ingles"
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
      Index           =   4
      Left            =   120
      TabIndex        =   178
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Control de Cambios"
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
      Index           =   3
      Left            =   120
      TabIndex        =   153
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
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
      Height          =   255
      Left            =   6720
      TabIndex        =   36
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label DesOperador 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   8280
      TabIndex        =   35
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblLabels 
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
      Index           =   2
      Left            =   8160
      TabIndex        =   32
      Top             =   360
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Version"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   31
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Descriprod 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   15
      Left            =   2040
      TabIndex        =   15
      Top             =   3360
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descripcion:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "PrgEspecifiUnifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstListaEspe As Recordset
Dim spListaEspe As String
Dim rstArticulo As Recordset
Dim spArticulo As String
Dim rstEnsayo As Recordset
Dim spEnsayo As String
Dim rstEspecificacionesUnifica As Recordset
Dim spEspecificacionesUnifica As String
Dim rstEspecificacionesUnificaII As Recordset
Dim spEspecificacionesUnificaII As String
Dim rstEspecificacionesUnificaIII As Recordset
Dim spEspecificacionesUnificaIII As String
Dim rstEspecificacionesUnificaVersion As Recordset
Dim spEspecificacionesUnificaVersion As String
Dim rstEspecificacionesUnificaVersionII As Recordset
Dim spEspecificacionesUnificaVersionII As String
Dim rstCertificadoMp As Recordset
Dim spCertificadoMp As String
Dim XParam As String
Dim EmpresaActual As String
Dim ZFecha As String
Dim ZVersion As String
Private WGraba As String
Dim ZOperador As String
Dim ZZOperador As String
Dim ZZVersion As String
Dim ZZFecha As String
Dim ZZProceso As Integer

Dim ZVector(10000) As String
Dim ZEnsayo(30) As String
Dim ZValor(30) As String
Dim ZDescri(30) As String
Dim ZDescriII(30) As String

Dim ZEnsayo1 As String
Dim ZEnsayo2 As String
Dim ZEnsayo3 As String
Dim ZEnsayo4 As String
Dim ZEnsayo5 As String
Dim ZEnsayo6 As String
Dim ZEnsayo7 As String
Dim ZEnsayo8 As String
Dim ZEnsayo9 As String
Dim ZEnsayo10 As String
Dim ZEnsayo11 As String
Dim ZEnsayo12 As String
Dim ZEnsayo13 As String
Dim ZEnsayo14 As String
Dim ZEnsayo15 As String
Dim ZEnsayo16 As String
Dim ZEnsayo17 As String
Dim ZEnsayo18 As String
Dim ZEnsayo19 As String
Dim ZEnsayo20 As String
Dim ZEnsayo21 As String
Dim ZEnsayo22 As String
Dim ZEnsayo23 As String
Dim ZEnsayo24 As String
Dim ZEnsayo25 As String
Dim ZEnsayo26 As String
Dim ZEnsayo27 As String
Dim ZEnsayo28 As String
Dim ZEnsayo29 As String
Dim ZEnsayo30 As String

Dim ZValor1 As String
Dim ZValor2 As String
Dim ZValor3 As String
Dim ZValor4 As String
Dim ZValor5 As String
Dim ZValor6 As String
Dim ZValor7 As String
Dim ZValor8 As String
Dim ZValor9 As String
Dim ZValor10 As String
Dim ZValor11 As String
Dim ZValor12 As String
Dim ZValor13 As String
Dim ZValor14 As String
Dim ZValor15 As String
Dim ZValor16 As String
Dim ZValor17 As String
Dim ZValor18 As String
Dim ZValor19 As String
Dim ZValor20 As String
Dim ZValor21 As String
Dim ZValor22 As String
Dim ZValor23 As String
Dim ZValor24 As String
Dim ZValor25 As String
Dim ZValor26 As String
Dim ZValor27 As String
Dim ZValor28 As String
Dim ZValor29 As String
Dim ZValor30 As String

Dim ZIValor1 As String
Dim ZIValor2 As String
Dim ZIValor3 As String
Dim ZIValor4 As String
Dim ZIValor5 As String
Dim ZIValor6 As String
Dim ZIValor7 As String
Dim ZIValor8 As String
Dim ZIValor9 As String
Dim ZIValor10 As String
Dim ZIValor11 As String
Dim ZIValor12 As String
Dim ZIValor13 As String
Dim ZIValor14 As String
Dim ZIValor15 As String
Dim ZIValor16 As String
Dim ZIValor17 As String
Dim ZIValor18 As String
Dim ZIValor19 As String
Dim ZIValor20 As String
Dim ZIValor21 As String
Dim ZIValor22 As String
Dim ZIValor23 As String
Dim ZIValor24 As String
Dim ZIValor25 As String
Dim ZIValor26 As String
Dim ZIValor27 As String
Dim ZIValor28 As String
Dim ZIValor29 As String
Dim ZIValor30 As String

Private Sub Imprime_Datos()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Erase ZEnsayo
        
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        Ensayo1.Text = rstEspecificacionesUnifica!Ensayo1
        Ensayo2.Text = rstEspecificacionesUnifica!Ensayo2
        Ensayo3.Text = rstEspecificacionesUnifica!Ensayo3
        Ensayo4.Text = rstEspecificacionesUnifica!Ensayo4
        Ensayo5.Text = rstEspecificacionesUnifica!Ensayo5
        Ensayo6.Text = rstEspecificacionesUnifica!Ensayo6
        Ensayo7.Text = rstEspecificacionesUnifica!Ensayo7
        Ensayo8.Text = rstEspecificacionesUnifica!Ensayo8
        Ensayo9.Text = rstEspecificacionesUnifica!Ensayo9
        Ensayo10.Text = rstEspecificacionesUnifica!Ensayo10
        Ensayo11.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        Ensayo12.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        Ensayo13.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        Ensayo14.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        Ensayo15.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        Ensayo16.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        Ensayo17.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        Ensayo18.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        Ensayo19.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        Ensayo20.Text = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        Valor1.Text = rstEspecificacionesUnifica!Valor1
        valor2.Text = rstEspecificacionesUnifica!valor2
        Valor3.Text = rstEspecificacionesUnifica!Valor3
        valor4.Text = rstEspecificacionesUnifica!valor4
        valor5.Text = rstEspecificacionesUnifica!valor5
        valor6.Text = rstEspecificacionesUnifica!valor6
        valor7.Text = rstEspecificacionesUnifica!valor7
        valor8.Text = rstEspecificacionesUnifica!valor8
        valor9.Text = rstEspecificacionesUnifica!valor9
        valor10.Text = rstEspecificacionesUnifica!valor10
        Valor11.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
        Valor12.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
        Valor13.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
        Valor14.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
        Valor15.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
        Valor16.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
        Valor17.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
        Valor18.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
        Valor19.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
        Valor20.Text = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
        
        Desde1.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde1), "", rstEspecificacionesUnifica!Desde1)
        Desde2.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde2), "", rstEspecificacionesUnifica!Desde2)
        Desde3.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde3), "", rstEspecificacionesUnifica!Desde3)
        Desde4.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde4), "", rstEspecificacionesUnifica!Desde4)
        Desde5.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde5), "", rstEspecificacionesUnifica!Desde5)
        Desde6.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde6), "", rstEspecificacionesUnifica!Desde6)
        Desde7.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde7), "", rstEspecificacionesUnifica!Desde7)
        Desde8.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde8), "", rstEspecificacionesUnifica!Desde8)
        Desde9.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde9), "", rstEspecificacionesUnifica!Desde9)
        Desde10.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde10), "", rstEspecificacionesUnifica!Desde10)
        Desde11.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde11), "", rstEspecificacionesUnifica!Desde11)
        Desde12.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde12), "", rstEspecificacionesUnifica!Desde12)
        Desde13.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde13), "", rstEspecificacionesUnifica!Desde13)
        Desde14.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde14), "", rstEspecificacionesUnifica!Desde14)
        Desde15.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde15), "", rstEspecificacionesUnifica!Desde15)
        Desde16.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde16), "", rstEspecificacionesUnifica!Desde16)
        Desde17.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde17), "", rstEspecificacionesUnifica!Desde17)
        Desde18.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde18), "", rstEspecificacionesUnifica!Desde18)
        Desde19.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde19), "", rstEspecificacionesUnifica!Desde19)
        Desde20.Text = IIf(IsNull(rstEspecificacionesUnifica!Desde20), "", rstEspecificacionesUnifica!Desde20)
        
        Hasta1.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta1), "", rstEspecificacionesUnifica!Hasta1)
        Hasta2.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta2), "", rstEspecificacionesUnifica!Hasta2)
        Hasta3.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta3), "", rstEspecificacionesUnifica!Hasta3)
        Hasta4.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta4), "", rstEspecificacionesUnifica!Hasta4)
        Hasta5.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta5), "", rstEspecificacionesUnifica!Hasta5)
        Hasta6.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta6), "", rstEspecificacionesUnifica!Hasta6)
        Hasta7.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta7), "", rstEspecificacionesUnifica!Hasta7)
        Hasta8.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta8), "", rstEspecificacionesUnifica!Hasta8)
        Hasta9.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta9), "", rstEspecificacionesUnifica!Hasta9)
        Hasta10.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta10), "", rstEspecificacionesUnifica!Hasta10)
        Hasta11.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta11), "", rstEspecificacionesUnifica!Hasta11)
        Hasta12.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta12), "", rstEspecificacionesUnifica!Hasta12)
        Hasta13.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta13), "", rstEspecificacionesUnifica!Hasta13)
        Hasta14.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta14), "", rstEspecificacionesUnifica!Hasta14)
        Hasta15.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta15), "", rstEspecificacionesUnifica!Hasta15)
        Hasta16.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta16), "", rstEspecificacionesUnifica!Hasta16)
        Hasta17.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta17), "", rstEspecificacionesUnifica!Hasta17)
        Hasta18.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta18), "", rstEspecificacionesUnifica!Hasta18)
        Hasta19.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta19), "", rstEspecificacionesUnifica!Hasta19)
        Hasta20.Text = IIf(IsNull(rstEspecificacionesUnifica!Hasta20), "", rstEspecificacionesUnifica!Hasta20)
       
        Valor1.Text = Trim(Valor1.Text)
        valor2.Text = Trim(valor2.Text)
        Valor3.Text = Trim(Valor3.Text)
        valor4.Text = Trim(valor4.Text)
        valor5.Text = Trim(valor5.Text)
        valor6.Text = Trim(valor6.Text)
        valor7.Text = Trim(valor7.Text)
        valor8.Text = Trim(valor8.Text)
        valor9.Text = Trim(valor9.Text)
        valor10.Text = Trim(valor10.Text)
        Valor11.Text = Trim(Valor11.Text)
        Valor12.Text = Trim(Valor12.Text)
        Valor13.Text = Trim(Valor13.Text)
        Valor14.Text = Trim(Valor14.Text)
        Valor15.Text = Trim(Valor15.Text)
        Valor16.Text = Trim(Valor16.Text)
        Valor17.Text = Trim(Valor17.Text)
        Valor18.Text = Trim(Valor18.Text)
        Valor19.Text = Trim(Valor19.Text)
        Valor20.Text = Trim(Valor20.Text)
        
        Desde1.Text = Trim(Desde1.Text)
        Desde2.Text = Trim(Desde2.Text)
        Desde3.Text = Trim(Desde3.Text)
        Desde4.Text = Trim(Desde4.Text)
        Desde5.Text = Trim(Desde5.Text)
        Desde6.Text = Trim(Desde6.Text)
        Desde7.Text = Trim(Desde7.Text)
        Desde8.Text = Trim(Desde8.Text)
        Desde9.Text = Trim(Desde9.Text)
        Desde10.Text = Trim(Desde10.Text)
        Desde11.Text = Trim(Desde11.Text)
        Desde12.Text = Trim(Desde12.Text)
        Desde13.Text = Trim(Desde13.Text)
        Desde14.Text = Trim(Desde14.Text)
        Desde15.Text = Trim(Desde15.Text)
        Desde16.Text = Trim(Desde16.Text)
        Desde17.Text = Trim(Desde17.Text)
        Desde18.Text = Trim(Desde18.Text)
        Desde19.Text = Trim(Desde19.Text)
        Desde20.Text = Trim(Desde20.Text)
        
        Hasta1.Text = Trim(Hasta1.Text)
        Hasta2.Text = Trim(Hasta2.Text)
        Hasta3.Text = Trim(Hasta3.Text)
        Hasta4.Text = Trim(Hasta4.Text)
        Hasta5.Text = Trim(Hasta5.Text)
        Hasta6.Text = Trim(Hasta6.Text)
        Hasta7.Text = Trim(Hasta7.Text)
        Hasta8.Text = Trim(Hasta8.Text)
        Hasta9.Text = Trim(Hasta9.Text)
        Hasta10.Text = Trim(Hasta10.Text)
        Hasta11.Text = Trim(Hasta11.Text)
        Hasta12.Text = Trim(Hasta12.Text)
        Hasta13.Text = Trim(Hasta13.Text)
        Hasta14.Text = Trim(Hasta14.Text)
        Hasta15.Text = Trim(Hasta15.Text)
        Hasta16.Text = Trim(Hasta16.Text)
        Hasta17.Text = Trim(Hasta17.Text)
        Hasta18.Text = Trim(Hasta18.Text)
        Hasta19.Text = Trim(Hasta19.Text)
        Hasta20.Text = Trim(Hasta20.Text)
        
        Version.Text = rstEspecificacionesUnifica!Version
        fecha.Text = rstEspecificacionesUnifica!fecha
        ZOperador = IIf(IsNull(rstEspecificacionesUnifica!Operador), "O", rstEspecificacionesUnifica!Operador)
        
        ControlCambio.Text = IIf(IsNull(rstEspecificacionesUnifica!ControlCambio), "", rstEspecificacionesUnifica!ControlCambio)
        
        rstEspecificacionesUnifica.Close
                        
    End If
    
    IValor1.Text = ""
    IValor2.Text = ""
    IValor3.Text = ""
    IValor4.Text = ""
    IValor5.Text = ""
    IValor6.Text = ""
    IValor7.Text = ""
    IValor8.Text = ""
    IValor9.Text = ""
    IValor10.Text = ""
    IValor11.Text = ""
    IValor12.Text = ""
    IValor13.Text = ""
    IValor14.Text = ""
    IValor15.Text = ""
    IValor16.Text = ""
    IValor17.Text = ""
    IValor18.Text = ""
    IValor19.Text = ""
    IValor20.Text = ""
    IValor21.Text = ""
    IValor22.Text = ""
    IValor23.Text = ""
    IValor24.Text = ""
    IValor25.Text = ""
    IValor26.Text = ""
    IValor27.Text = ""
    IValor28.Text = ""
    IValor29.Text = ""
    IValor30.Text = ""

    Ensayo21.Text = ""
    
    Ensayo22.Text = ""
    Ensayo23.Text = ""
    Ensayo24.Text = ""
    Ensayo25.Text = ""
    Ensayo26.Text = ""
    Ensayo27.Text = ""
    Ensayo28.Text = ""
    Ensayo29.Text = ""
    Ensayo30.Text = ""
        
    Valor21.Text = ""
    Valor22.Text = ""
    Valor23.Text = ""
    Valor24.Text = ""
    Valor25.Text = ""
    Valor26.Text = ""
    Valor27.Text = ""
    Valor28.Text = ""
    Valor29.Text = ""
    Valor30.Text = ""
        
    Desde21.Text = ""
    Desde22.Text = ""
    Desde23.Text = ""
    Desde24.Text = ""
    Desde25.Text = ""
    Desde26.Text = ""
    Desde27.Text = ""
    Desde28.Text = ""
    Desde29.Text = ""
    Desde30.Text = ""
       
    Hasta21.Text = ""
    Hasta22.Text = ""
    Hasta23.Text = ""
    Hasta24.Text = ""
    Hasta25.Text = ""
    Hasta26.Text = ""
    Hasta27.Text = ""
    Hasta28.Text = ""
    Hasta29.Text = ""
    Hasta30.Text = ""
       
    DescripcionIngles.Text = ""
    Cas.Text = ""
    
    DescripcionIngles.Text = ""
    Cas.Text = ""
    
    Sql1 = "Select EspecificacionesUnificaII.IValor1, EspecificacionesUnificaII.IValor2, EspecificacionesUnificaII.IValor3, EspecificacionesUnificaII.IValor4, EspecificacionesUnificaII.IValor5, EspecificacionesUnificaII.IValor6, EspecificacionesUnificaII.IValor7, EspecificacionesUnificaII.IValor8, EspecificacionesUnificaII.IValor9, EspecificacionesUnificaII.IValor10, "
    Sql2 = "       EspecificacionesUnificaII.IValor11, EspecificacionesUnificaII.IValor12, EspecificacionesUnificaII.IValor13, EspecificacionesUnificaII.IValor14, EspecificacionesUnificaII.IValor15, EspecificacionesUnificaII.IValor16, EspecificacionesUnificaII.IValor17, EspecificacionesUnificaII.IValor18, EspecificacionesUnificaII.IValor19, EspecificacionesUnificaII.IValor20, "
    Sql3 = "       EspecificacionesUnificaII.DescripcionIngles, EspecificacionesUnificaII.cas"
    Sql4 = " FROM EspecificacionesUnificaII"
    Sql5 = " Where EspecificacionesUnificaII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaII.RecordCount > 0 Then
        
        IValor1.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor1), "", rstEspecificacionesUnificaII!IValor1)
        IValor2.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor2), "", rstEspecificacionesUnificaII!IValor2)
        IValor3.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor3), "", rstEspecificacionesUnificaII!IValor3)
        IValor4.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor4), "", rstEspecificacionesUnificaII!IValor4)
        IValor5.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor5), "", rstEspecificacionesUnificaII!IValor5)
        IValor6.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor6), "", rstEspecificacionesUnificaII!IValor6)
        IValor7.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor7), "", rstEspecificacionesUnificaII!IValor7)
        IValor8.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor8), "", rstEspecificacionesUnificaII!IValor8)
        IValor9.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor9), "", rstEspecificacionesUnificaII!IValor9)
        IValor10.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor10), "", rstEspecificacionesUnificaII!IValor10)
        IValor11.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor11), "", rstEspecificacionesUnificaII!IValor11)
        IValor12.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor12), "", rstEspecificacionesUnificaII!IValor12)
        IValor13.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor13), "", rstEspecificacionesUnificaII!IValor13)
        IValor14.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor14), "", rstEspecificacionesUnificaII!IValor14)
        IValor15.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor15), "", rstEspecificacionesUnificaII!IValor15)
        IValor16.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor16), "", rstEspecificacionesUnificaII!IValor16)
        IValor17.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor17), "", rstEspecificacionesUnificaII!IValor17)
        IValor18.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor18), "", rstEspecificacionesUnificaII!IValor18)
        IValor19.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor19), "", rstEspecificacionesUnificaII!IValor19)
        IValor20.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor20), "", rstEspecificacionesUnificaII!IValor20)
       
        DescripcionIngles.Text = IIf(IsNull(rstEspecificacionesUnificaII!DescripcionIngles), "", rstEspecificacionesUnificaII!DescripcionIngles)
        Cas.Text = IIf(IsNull(rstEspecificacionesUnificaII!Cas), "", rstEspecificacionesUnificaII!Cas)
       
        DescripcionIngles.Text = Trim(DescripcionIngles.Text)
        Cas.Text = Trim(Cas.Text)
       
        IValor1.Text = Trim(IValor1.Text)
        IValor2.Text = Trim(IValor2.Text)
        IValor3.Text = Trim(IValor3.Text)
        IValor4.Text = Trim(IValor4.Text)
        IValor5.Text = Trim(IValor5.Text)
        IValor6.Text = Trim(IValor6.Text)
        IValor7.Text = Trim(IValor7.Text)
        IValor8.Text = Trim(IValor8.Text)
        IValor9.Text = Trim(IValor9.Text)
        IValor10.Text = Trim(IValor10.Text)
        IValor11.Text = Trim(IValor11.Text)
        IValor12.Text = Trim(IValor12.Text)
        IValor13.Text = Trim(IValor13.Text)
        IValor14.Text = Trim(IValor14.Text)
        IValor15.Text = Trim(IValor15.Text)
        IValor16.Text = Trim(IValor16.Text)
        IValor17.Text = Trim(IValor17.Text)
        IValor18.Text = Trim(IValor18.Text)
        IValor19.Text = Trim(IValor19.Text)
        IValor20.Text = Trim(IValor20.Text)
        
        rstEspecificacionesUnificaII.Close
                        
    End If
    
    
    
    Sql1 = "Select EspecificacionesUnificaII.IValor21, EspecificacionesUnificaII.IValor22, EspecificacionesUnificaII.IValor23, EspecificacionesUnificaII.IValor24, EspecificacionesUnificaII.IValor25, EspecificacionesUnificaII.IValor26, EspecificacionesUnificaII.IValor27, EspecificacionesUnificaII.IValor28, EspecificacionesUnificaII.IValor29, EspecificacionesUnificaII.IValor30 "
    Sql2 = " FROM EspecificacionesUnificaII"
    Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaII.RecordCount > 0 Then
        
        IValor21.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor21), "", rstEspecificacionesUnificaII!IValor21)
        IValor22.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor22), "", rstEspecificacionesUnificaII!IValor22)
        IValor23.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor23), "", rstEspecificacionesUnificaII!IValor23)
        IValor24.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor24), "", rstEspecificacionesUnificaII!IValor24)
        IValor25.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor25), "", rstEspecificacionesUnificaII!IValor25)
        IValor26.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor26), "", rstEspecificacionesUnificaII!IValor26)
        IValor27.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor27), "", rstEspecificacionesUnificaII!IValor27)
        IValor28.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor28), "", rstEspecificacionesUnificaII!IValor28)
        IValor29.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor29), "", rstEspecificacionesUnificaII!IValor29)
        IValor30.Text = IIf(IsNull(rstEspecificacionesUnificaII!IValor30), "", rstEspecificacionesUnificaII!IValor30)
       
        IValor21.Text = Trim(IValor21.Text)
        IValor22.Text = Trim(IValor22.Text)
        IValor23.Text = Trim(IValor23.Text)
        IValor24.Text = Trim(IValor24.Text)
        IValor25.Text = Trim(IValor25.Text)
        IValor26.Text = Trim(IValor26.Text)
        IValor27.Text = Trim(IValor27.Text)
        IValor28.Text = Trim(IValor28.Text)
        IValor29.Text = Trim(IValor29.Text)
        IValor30.Text = Trim(IValor30.Text)
        
        rstEspecificacionesUnificaII.Close
                        
    End If
    
    
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
       
        Ensayo21.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        Ensayo22.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        Ensayo23.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        Ensayo24.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        Ensayo25.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        Ensayo26.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        Ensayo27.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        Ensayo28.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        Ensayo29.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        Ensayo30.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
       
        Valor21.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
        Valor22.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
        Valor23.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
        Valor24.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
        Valor25.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
        Valor26.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
        Valor27.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
        Valor28.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
        Valor29.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
        Valor30.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
        
        Valor21.Text = Trim(Valor21.Text)
        Valor22.Text = Trim(Valor22.Text)
        Valor23.Text = Trim(Valor23.Text)
        Valor24.Text = Trim(Valor24.Text)
        Valor25.Text = Trim(Valor25.Text)
        Valor26.Text = Trim(Valor26.Text)
        Valor27.Text = Trim(Valor27.Text)
        Valor28.Text = Trim(Valor28.Text)
        Valor29.Text = Trim(Valor29.Text)
        Valor30.Text = Trim(Valor30.Text)
        
        Desde21.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde21), "", rstEspecificacionesUnificaIII!Desde21)
        Desde22.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde22), "", rstEspecificacionesUnificaIII!Desde22)
        Desde23.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde23), "", rstEspecificacionesUnificaIII!Desde23)
        Desde24.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde24), "", rstEspecificacionesUnificaIII!Desde24)
        Desde25.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde25), "", rstEspecificacionesUnificaIII!Desde25)
        Desde26.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde26), "", rstEspecificacionesUnificaIII!Desde26)
        Desde27.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde27), "", rstEspecificacionesUnificaIII!Desde27)
        Desde28.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde28), "", rstEspecificacionesUnificaIII!Desde28)
        Desde29.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde29), "", rstEspecificacionesUnificaIII!Desde29)
        Desde30.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Desde30), "", rstEspecificacionesUnificaIII!Desde30)
        
        Desde20.Text = Trim(Desde20.Text)
        Desde21.Text = Trim(Desde21.Text)
        Desde22.Text = Trim(Desde22.Text)
        Desde23.Text = Trim(Desde23.Text)
        Desde24.Text = Trim(Desde24.Text)
        Desde25.Text = Trim(Desde25.Text)
        Desde26.Text = Trim(Desde26.Text)
        Desde27.Text = Trim(Desde27.Text)
        Desde28.Text = Trim(Desde28.Text)
        Desde29.Text = Trim(Desde29.Text)
        Desde30.Text = Trim(Desde30.Text)
       
        Hasta21.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta21), "", rstEspecificacionesUnificaIII!Hasta21)
        Hasta22.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta22), "", rstEspecificacionesUnificaIII!Hasta22)
        Hasta23.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta23), "", rstEspecificacionesUnificaIII!Hasta23)
        Hasta24.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta24), "", rstEspecificacionesUnificaIII!Hasta24)
        Hasta25.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta25), "", rstEspecificacionesUnificaIII!Hasta25)
        Hasta26.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta26), "", rstEspecificacionesUnificaIII!Hasta26)
        Hasta27.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta27), "", rstEspecificacionesUnificaIII!Hasta27)
        Hasta28.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta28), "", rstEspecificacionesUnificaIII!Hasta28)
        Hasta29.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta29), "", rstEspecificacionesUnificaIII!Hasta29)
        Hasta30.Text = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta30), "", rstEspecificacionesUnificaIII!Hasta30)
        
        Hasta21.Text = Trim(Hasta21.Text)
        Hasta22.Text = Trim(Hasta22.Text)
        Hasta23.Text = Trim(Hasta23.Text)
        Hasta24.Text = Trim(Hasta24.Text)
        Hasta25.Text = Trim(Hasta25.Text)
        Hasta26.Text = Trim(Hasta26.Text)
        Hasta27.Text = Trim(Hasta27.Text)
        Hasta28.Text = Trim(Hasta28.Text)
        Hasta29.Text = Trim(Hasta29.Text)
        Hasta30.Text = Trim(Hasta30.Text)
        
        rstEspecificacionesUnificaIII.Close
                        
    End If
    
    For Cicla = 1 To 30
        ZZDescri = ""
        If Val(ZEnsayo(Cicla)) <> 0 Then
            spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDescri = rstEnsayo!Descripcion
                rstEnsayo.Close
            End If
        End If
        Select Case Cicla
            Case 1
                Descri1.Caption = ZZDescri
            Case 2
                descri2.Caption = ZZDescri
            Case 3
                Descri3.Caption = ZZDescri
            Case 4
                Descri4.Caption = ZZDescri
            Case 5
                Descri5.Caption = ZZDescri
            Case 6
                Descri6.Caption = ZZDescri
            Case 7
                Descri7.Caption = ZZDescri
            Case 8
                Descri8.Caption = ZZDescri
            Case 9
                Descri9.Caption = ZZDescri
            Case 10
                Descri10.Caption = ZZDescri
            Case 11
                Descri11.Caption = ZZDescri
            Case 12
                Descri12.Caption = ZZDescri
            Case 13
                Descri13.Caption = ZZDescri
            Case 14
                Descri14.Caption = ZZDescri
            Case 15
                Descri15.Caption = ZZDescri
            Case 16
                Descri16.Caption = ZZDescri
            Case 17
                Descri17.Caption = ZZDescri
            Case 18
                Descri18.Caption = ZZDescri
            Case 19
                Descri19.Caption = ZZDescri
            Case 20
                Descri20.Caption = ZZDescri
            Case 21
                Descri21.Caption = ZZDescri
            Case 22
                Descri22.Caption = ZZDescri
            Case 23
                Descri23.Caption = ZZDescri
            Case 24
                Descri24.Caption = ZZDescri
            Case 25
                Descri25.Caption = ZZDescri
            Case 26
                Descri26.Caption = ZZDescri
            Case 27
                Descri27.Caption = ZZDescri
            Case 28
                Descri28.Caption = ZZDescri
            Case 29
                Descri29.Caption = ZZDescri
            Case 30
                Descri30.Caption = ZZDescri
            Case Else
        End Select
                
    Next Cicla
    
    DesOperador.Caption = ""
    If Val(ZOperador) <> 0 Then
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Operador = " + "'" + ZOperador + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            DesOperador.Caption = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
            rstOperador.Close
        End If
    End If
    
    Call Conecta_Empresa
        
    spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
    Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
    If rstArticulo.RecordCount > 0 Then
        Descriprod.Caption = rstArticulo!Descripcion
        rstArticulo.Close
    End If

End Sub

Private Sub Acepta_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Erase ZVector
    ZLugar = 0
    
    Desde.Text = UCase(Desde.Text)
    Hasta.Text = UCase(Hasta.Text)
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = ZSql
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        With rstEspecificacionesUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstEspecificacionesUnifica!Producto >= Desde.Text And rstEspecificacionesUnifica!Producto <= Hasta.Text Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = rstEspecificacionesUnifica!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecificacionesUnifica.Close
    End If
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
        
        Erase ZEnsayo
        Erase ZValor
        Erase ZDescri
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecificacionesUnifica"
        ZSql = ZSql + " Where EspecificacionesUnifica.Producto = " + "'" + ZCodigo + "'"
        spEspecificacionesUnifica = ZSql
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnifica.RecordCount > 0 Then
    
            ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
            ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
            ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
            ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
            ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
            ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
            ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
            ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
            ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
            ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
            ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
            ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
            ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
            ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
            ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
            ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
            ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
            ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
            ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
            ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
            
            ZValor(1) = rstEspecificacionesUnifica!Valor1
            ZValor(2) = rstEspecificacionesUnifica!valor2
            ZValor(3) = rstEspecificacionesUnifica!Valor3
            ZValor(4) = rstEspecificacionesUnifica!valor4
            ZValor(5) = rstEspecificacionesUnifica!valor5
            ZValor(6) = rstEspecificacionesUnifica!valor6
            ZValor(7) = rstEspecificacionesUnifica!valor7
            ZValor(8) = rstEspecificacionesUnifica!valor8
            ZValor(9) = rstEspecificacionesUnifica!valor9
            ZValor(10) = rstEspecificacionesUnifica!valor10
            ZValor(11) = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
            ZValor(12) = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
            ZValor(13) = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
            ZValor(14) = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
            ZValor(15) = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
            ZValor(16) = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
            ZValor(17) = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
            ZValor(18) = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
            ZValor(19) = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
            ZValor(20) = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
            
            ZZOperador = IIf(IsNull(rstEspecificacionesUnifica!Operador), "O", rstEspecificacionesUnifica!Operador)
            ZZVersion = rstEspecificacionesUnifica!Version
            ZZFecha = rstEspecificacionesUnifica!fecha
        
            rstEspecificacionesUnifica.Close
                        
        End If
    
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecificacionesUnificaIII"
        ZSql = ZSql + " Where EspecificacionesUnificaIII.Producto = " + "'" + ZCodigo + "'"
        spEspecificacionesUnificaIII = ZSql
        Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
            ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
            
            ZValor(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
            ZValor(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
            ZValor(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
            ZValor(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
            ZValor(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
            ZValor(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
            ZValor(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
            ZValor(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
            ZValor(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
            ZValor(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
        
            rstEspecificacionesUnificaIII.Close
                        
        End If
    
        For Cicla = 1 To 30
            ZDescri(Cicla) = ""
            If Val(ZEnsayo(Cicla)) <> 0 Then
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    ZDescri(Cicla) = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
            End If
        Next Cicla
    
        spArticulo = "ConsultaArticulo " + "'" + ZCodigo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        ZZDesOperador = ""
        If Val(ZZOperador) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.Operador = " + "'" + ZZOperador + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZZDesOperador = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
                rstOperador.Close
            End If
        End If
        
        
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaEspe ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion,"
        ZSql = ZSql + "Codigo1,"
        ZSql = ZSql + "Codigo2,"
        ZSql = ZSql + "Codigo3,"
        ZSql = ZSql + "Codigo4,"
        ZSql = ZSql + "Codigo5,"
        ZSql = ZSql + "Codigo6,"
        ZSql = ZSql + "Codigo7,"
        ZSql = ZSql + "Codigo8,"
        ZSql = ZSql + "Codigo9,"
        ZSql = ZSql + "Codigo10,"
        ZSql = ZSql + "Codigo11,"
        ZSql = ZSql + "Codigo12,"
        ZSql = ZSql + "Codigo13,"
        ZSql = ZSql + "Codigo14,"
        ZSql = ZSql + "Codigo15,"
        ZSql = ZSql + "Codigo16,"
        ZSql = ZSql + "Codigo17,"
        ZSql = ZSql + "Codigo18,"
        ZSql = ZSql + "Codigo19,"
        ZSql = ZSql + "Codigo20,"
        ZSql = ZSql + "Codigo21,"
        ZSql = ZSql + "Codigo22,"
        ZSql = ZSql + "Codigo23,"
        ZSql = ZSql + "Codigo24,"
        ZSql = ZSql + "Codigo25,"
        ZSql = ZSql + "Codigo26,"
        ZSql = ZSql + "Codigo27,"
        ZSql = ZSql + "Codigo28,"
        ZSql = ZSql + "Codigo29,"
        ZSql = ZSql + "Codigo30,"
        ZSql = ZSql + "Descri1,"
        ZSql = ZSql + "Descri2,"
        ZSql = ZSql + "Descri3,"
        ZSql = ZSql + "Descri4,"
        ZSql = ZSql + "Descri5,"
        ZSql = ZSql + "Descri6,"
        ZSql = ZSql + "Descri7,"
        ZSql = ZSql + "Descri8,"
        ZSql = ZSql + "Descri9,"
        ZSql = ZSql + "Descri10,"
        ZSql = ZSql + "Descri11,"
        ZSql = ZSql + "Descri12,"
        ZSql = ZSql + "Descri13,"
        ZSql = ZSql + "Descri14,"
        ZSql = ZSql + "Descri15,"
        ZSql = ZSql + "Descri16,"
        ZSql = ZSql + "Descri17,"
        ZSql = ZSql + "Descri18,"
        ZSql = ZSql + "Descri19,"
        ZSql = ZSql + "Descri20,"
        ZSql = ZSql + "Descri21,"
        ZSql = ZSql + "Descri22,"
        ZSql = ZSql + "Descri23,"
        ZSql = ZSql + "Descri24,"
        ZSql = ZSql + "Descri25,"
        ZSql = ZSql + "Descri26,"
        ZSql = ZSql + "Descri27,"
        ZSql = ZSql + "Descri28,"
        ZSql = ZSql + "Descri29,"
        ZSql = ZSql + "Descri30,"
        ZSql = ZSql + "Valor1,"
        ZSql = ZSql + "Valor2,"
        ZSql = ZSql + "Valor3,"
        ZSql = ZSql + "Valor4,"
        ZSql = ZSql + "Valor5,"
        ZSql = ZSql + "Valor6,"
        ZSql = ZSql + "Valor7,"
        ZSql = ZSql + "Valor8,"
        ZSql = ZSql + "Valor9,"
        ZSql = ZSql + "Valor10,"
        ZSql = ZSql + "ZValor1,"
        ZSql = ZSql + "ZValor2,"
        ZSql = ZSql + "ZValor3,"
        ZSql = ZSql + "ZValor4,"
        ZSql = ZSql + "ZValor5,"
        ZSql = ZSql + "ZValor6,"
        ZSql = ZSql + "ZValor7,"
        ZSql = ZSql + "ZValor8,"
        ZSql = ZSql + "ZValor9,"
        ZSql = ZSql + "ZValor10,"
        ZSql = ZSql + "ZValor121,"
        ZSql = ZSql + "ZValor122,"
        ZSql = ZSql + "ZValor123,"
        ZSql = ZSql + "ZValor124,"
        ZSql = ZSql + "ZValor125,"
        ZSql = ZSql + "ZValor126,"
        ZSql = ZSql + "ZValor127,"
        ZSql = ZSql + "ZValor128,"
        ZSql = ZSql + "ZValor129,"
        ZSql = ZSql + "ZValor130,"
        ZSql = ZSql + "Version ,"
        ZSql = ZSql + "Responsable,"
        ZSql = ZSql + "Fecha )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZEnsayo(1) + "',"
        ZSql = ZSql + "'" + ZEnsayo(2) + "',"
        ZSql = ZSql + "'" + ZEnsayo(3) + "',"
        ZSql = ZSql + "'" + ZEnsayo(4) + "',"
        ZSql = ZSql + "'" + ZEnsayo(5) + "',"
        ZSql = ZSql + "'" + ZEnsayo(6) + "',"
        ZSql = ZSql + "'" + ZEnsayo(7) + "',"
        ZSql = ZSql + "'" + ZEnsayo(8) + "',"
        ZSql = ZSql + "'" + ZEnsayo(9) + "',"
        ZSql = ZSql + "'" + ZEnsayo(10) + "',"
        ZSql = ZSql + "'" + ZEnsayo(11) + "',"
        ZSql = ZSql + "'" + ZEnsayo(12) + "',"
        ZSql = ZSql + "'" + ZEnsayo(13) + "',"
        ZSql = ZSql + "'" + ZEnsayo(14) + "',"
        ZSql = ZSql + "'" + ZEnsayo(15) + "',"
        ZSql = ZSql + "'" + ZEnsayo(16) + "',"
        ZSql = ZSql + "'" + ZEnsayo(17) + "',"
        ZSql = ZSql + "'" + ZEnsayo(18) + "',"
        ZSql = ZSql + "'" + ZEnsayo(19) + "',"
        ZSql = ZSql + "'" + ZEnsayo(20) + "',"
        ZSql = ZSql + "'" + ZEnsayo(21) + "',"
        ZSql = ZSql + "'" + ZEnsayo(22) + "',"
        ZSql = ZSql + "'" + ZEnsayo(23) + "',"
        ZSql = ZSql + "'" + ZEnsayo(24) + "',"
        ZSql = ZSql + "'" + ZEnsayo(25) + "',"
        ZSql = ZSql + "'" + ZEnsayo(26) + "',"
        ZSql = ZSql + "'" + ZEnsayo(27) + "',"
        ZSql = ZSql + "'" + ZEnsayo(28) + "',"
        ZSql = ZSql + "'" + ZEnsayo(29) + "',"
        ZSql = ZSql + "'" + ZEnsayo(30) + "',"
        ZSql = ZSql + "'" + ZDescri(1) + "',"
        ZSql = ZSql + "'" + ZDescri(2) + "',"
        ZSql = ZSql + "'" + ZDescri(3) + "',"
        ZSql = ZSql + "'" + ZDescri(4) + "',"
        ZSql = ZSql + "'" + ZDescri(5) + "',"
        ZSql = ZSql + "'" + ZDescri(6) + "',"
        ZSql = ZSql + "'" + ZDescri(7) + "',"
        ZSql = ZSql + "'" + ZDescri(8) + "',"
        ZSql = ZSql + "'" + ZDescri(9) + "',"
        ZSql = ZSql + "'" + ZDescri(10) + "',"
        ZSql = ZSql + "'" + ZDescri(11) + "',"
        ZSql = ZSql + "'" + ZDescri(12) + "',"
        ZSql = ZSql + "'" + ZDescri(13) + "',"
        ZSql = ZSql + "'" + ZDescri(14) + "',"
        ZSql = ZSql + "'" + ZDescri(15) + "',"
        ZSql = ZSql + "'" + ZDescri(16) + "',"
        ZSql = ZSql + "'" + ZDescri(17) + "',"
        ZSql = ZSql + "'" + ZDescri(18) + "',"
        ZSql = ZSql + "'" + ZDescri(19) + "',"
        ZSql = ZSql + "'" + ZDescri(20) + "',"
        ZSql = ZSql + "'" + ZDescri(21) + "',"
        ZSql = ZSql + "'" + ZDescri(22) + "',"
        ZSql = ZSql + "'" + ZDescri(23) + "',"
        ZSql = ZSql + "'" + ZDescri(24) + "',"
        ZSql = ZSql + "'" + ZDescri(25) + "',"
        ZSql = ZSql + "'" + ZDescri(26) + "',"
        ZSql = ZSql + "'" + ZDescri(27) + "',"
        ZSql = ZSql + "'" + ZDescri(28) + "',"
        ZSql = ZSql + "'" + ZDescri(29) + "',"
        ZSql = ZSql + "'" + ZDescri(30) + "',"
        ZSql = ZSql + "'" + ZValor(1) + "',"
        ZSql = ZSql + "'" + ZValor(2) + "',"
        ZSql = ZSql + "'" + ZValor(3) + "',"
        ZSql = ZSql + "'" + ZValor(4) + "',"
        ZSql = ZSql + "'" + ZValor(5) + "',"
        ZSql = ZSql + "'" + ZValor(6) + "',"
        ZSql = ZSql + "'" + ZValor(7) + "',"
        ZSql = ZSql + "'" + ZValor(8) + "',"
        ZSql = ZSql + "'" + ZValor(9) + "',"
        ZSql = ZSql + "'" + ZValor(10) + "',"
        ZSql = ZSql + "'" + ZValor(11) + "',"
        ZSql = ZSql + "'" + ZValor(12) + "',"
        ZSql = ZSql + "'" + ZValor(13) + "',"
        ZSql = ZSql + "'" + ZValor(14) + "',"
        ZSql = ZSql + "'" + ZValor(15) + "',"
        ZSql = ZSql + "'" + ZValor(16) + "',"
        ZSql = ZSql + "'" + ZValor(17) + "',"
        ZSql = ZSql + "'" + ZValor(18) + "',"
        ZSql = ZSql + "'" + ZValor(19) + "',"
        ZSql = ZSql + "'" + ZValor(20) + "',"
        ZSql = ZSql + "'" + ZValor(21) + "',"
        ZSql = ZSql + "'" + ZValor(22) + "',"
        ZSql = ZSql + "'" + ZValor(23) + "',"
        ZSql = ZSql + "'" + ZValor(24) + "',"
        ZSql = ZSql + "'" + ZValor(25) + "',"
        ZSql = ZSql + "'" + ZValor(26) + "',"
        ZSql = ZSql + "'" + ZValor(27) + "',"
        ZSql = ZSql + "'" + ZValor(28) + "',"
        ZSql = ZSql + "'" + ZValor(29) + "',"
        ZSql = ZSql + "'" + ZValor(30) + "',"
        ZSql = ZSql + "'" + ZZVersion + "',"
        ZSql = ZSql + "'" + ZZDesOperador + "',"
        ZSql = ZSql + "'" + ZZFecha + "')"
        
        spListaEspe = ZSql
        Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ListaEspe SET "
    ZSql = ZSql + "ControlCambio = " + "'" + ControlCambio.Text + "'"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height

    Rem lista.GroupSelectionFormula = "{EspecificacionesUnifica.Producto} in " + Chr$(34) + Desde.Text + Chr$(34) + " to " + Chr$(34) + Hasta.Text + Chr$(34)
    If ImpreListado.Value = True Then
        lista.Destination = 1
            Else
        lista.Destination = 0
    End If
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            lista.ReportFileName = "ListaEspe.rpt"
        Case Else
            lista.ReportFileName = "ListaEspePelli.rpt"
    End Select
    
    Rem Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, " _
    rem             + "ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, " _
    rem             + "ListaEspe.Codigo11, ListaEspe.Codigo12, ListaEspe.Codigo13, ListaEspe.Codigo14, ListaEspe.Codigo15, ListaEspe.Codigo16, ListaEspe.Codigo17, ListaEspe.Codigo18, ListaEspe.Codigo19, ListaEspe.Codigo20, " _
    rem             + "ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, " _
    rem             + "ListaEspe.Descri11, ListaEspe.Descri12, ListaEspe.Descri13, ListaEspe.Descri14, ListaEspe.Descri15, ListaEspe.Descri16, ListaEspe.Descri17, ListaEspe.Descri18, ListaEspe.Descri19, ListaEspe.Descri20, " _
    rem             + "ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, " _
    rem             + "ListaEspe.ZValor1, ListaEspe.ZValor2, ListaEspe.ZValor3, ListaEspe.ZValor4, ListaEspe.ZValor5, ListaEspe.ZValor6, ListaEspe.ZValor7, ListaEspe.ZValor8, ListaEspe.ZValor9, ListaEspe.ZValor10, " _
    rem             + "ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha, ListaEspe.ControlCambio  " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ListaEspe ListaEspe " _
    rem             + "Where " _
    rem             + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
    rem             + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    
    lista.SQLQuery = "SELECT * " _
                + "From " _
                + DSQ + ".dbo.ListaEspe ListaEspe " _
                + "Where " _
                + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
                + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    
    lista.Connect = Connect()
    
    lista.Action = 1
    Frame2.Visible = False
    
    Call Conecta_Empresa
    
End Sub

Private Sub ImprimeAutomatico()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    
    Erase ZVector
    ZLugar = 0
    
    Desde.Text = Codigo.Text
    Hasta.Text = Codigo.Text
    
    ZSql = "DELETE ListaEspe"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    ZSql = ""
    ZSql = ZSql + "Select *"
    ZSql = ZSql + " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = ZSql
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        With rstEspecificacionesUnifica
            .MoveFirst
            Do
                If .EOF = False Then
                    If rstEspecificacionesUnifica!Producto >= Desde.Text And rstEspecificacionesUnifica!Producto <= Hasta.Text Then
                        ZLugar = ZLugar + 1
                        ZVector(ZLugar) = rstEspecificacionesUnifica!Producto
                    End If
                    .MoveNext
                        Else
                    Exit Do
                End If
            Loop
        End With
        rstEspecificacionesUnifica.Close
    End If
    
    
    
    
    
    For Ciclo = 1 To ZLugar
    
        ZCodigo = ZVector(Ciclo)
        
        Erase ZEnsayo
        Erase ZValor
        Erase ZDescri
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecificacionesUnifica"
        ZSql = ZSql + " Where EspecificacionesUnifica.Producto = " + "'" + ZCodigo + "'"
        spEspecificacionesUnifica = ZSql
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnifica.RecordCount > 0 Then
    
            ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
            ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
            ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
            ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
            ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
            ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
            ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
            ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
            ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
            ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
            ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
            ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
            ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
            ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
            ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
            ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
            ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
            ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
            ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
            ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
            
            ZValor(1) = rstEspecificacionesUnifica!Valor1
            ZValor(2) = rstEspecificacionesUnifica!valor2
            ZValor(3) = rstEspecificacionesUnifica!Valor3
            ZValor(4) = rstEspecificacionesUnifica!valor4
            ZValor(5) = rstEspecificacionesUnifica!valor5
            ZValor(6) = rstEspecificacionesUnifica!valor6
            ZValor(7) = rstEspecificacionesUnifica!valor7
            ZValor(8) = rstEspecificacionesUnifica!valor8
            ZValor(9) = rstEspecificacionesUnifica!valor9
            ZValor(10) = rstEspecificacionesUnifica!valor10
            ZValor(11) = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
            ZValor(12) = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
            ZValor(13) = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
            ZValor(14) = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
            ZValor(15) = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
            ZValor(16) = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
            ZValor(17) = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
            ZValor(18) = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
            ZValor(19) = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
            ZValor(20) = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
            
            ZZOperador = IIf(IsNull(rstEspecificacionesUnifica!Operador), "O", rstEspecificacionesUnifica!Operador)
            ZZVersion = rstEspecificacionesUnifica!Version
            ZZFecha = rstEspecificacionesUnifica!fecha
        
            rstEspecificacionesUnifica.Close
                        
        End If
    
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM EspecificacionesUnificaIII"
        ZSql = ZSql + " Where EspecificacionesUnificaIII.Producto = " + "'" + ZCodigo + "'"
        spEspecificacionesUnificaIII = ZSql
        Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
            ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
            ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
            ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
            ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
            ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
            ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
            ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
            ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
            ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
            ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
            
            ZValor(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
            ZValor(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
            ZValor(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
            ZValor(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
            ZValor(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
            ZValor(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
            ZValor(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
            ZValor(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
            ZValor(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
            ZValor(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
        
            rstEspecificacionesUnificaIII.Close
                        
        End If
    
        For Cicla = 1 To 30
            If Val(ZEnsayo(Cicla)) <> 0 Then
                spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
                Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
                If rstEnsayo.RecordCount > 0 Then
                    ZDescri(Cicla) = rstEnsayo!Descripcion
                    rstEnsayo.Close
                End If
            End If
        Next Cicla
    
        spArticulo = "ConsultaArticulo " + "'" + ZCodigo + "'"
        Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
        If rstArticulo.RecordCount > 0 Then
            ZDescripcion = rstArticulo!Descripcion
            rstArticulo.Close
        End If
        
        ZZDesOperador = ""
        If Val(ZZOperador) <> 0 Then
            ZSql = ""
            ZSql = ZSql + "Select *"
            ZSql = ZSql + " FROM Operador"
            ZSql = ZSql + " Where Operador.Operador = " + "'" + ZZOperador + "'"
            spOperador = ZSql
            Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
            If rstOperador.RecordCount > 0 Then
                ZZDesOperador = IIf(IsNull(rstOperador!Descripcion), "", rstOperador!Descripcion)
                rstOperador.Close
            End If
        End If
            
        ZSql = ""
        ZSql = ZSql + "INSERT INTO ListaEspe ("
        ZSql = ZSql + "Codigo ,"
        ZSql = ZSql + "Descripcion,"
        ZSql = ZSql + "Codigo1,"
        ZSql = ZSql + "Codigo2,"
        ZSql = ZSql + "Codigo3,"
        ZSql = ZSql + "Codigo4,"
        ZSql = ZSql + "Codigo5,"
        ZSql = ZSql + "Codigo6,"
        ZSql = ZSql + "Codigo7,"
        ZSql = ZSql + "Codigo8,"
        ZSql = ZSql + "Codigo9,"
        ZSql = ZSql + "Codigo10,"
        ZSql = ZSql + "Codigo11,"
        ZSql = ZSql + "Codigo12,"
        ZSql = ZSql + "Codigo13,"
        ZSql = ZSql + "Codigo14,"
        ZSql = ZSql + "Codigo15,"
        ZSql = ZSql + "Codigo16,"
        ZSql = ZSql + "Codigo17,"
        ZSql = ZSql + "Codigo18,"
        ZSql = ZSql + "Codigo19,"
        ZSql = ZSql + "Codigo20,"
        ZSql = ZSql + "Codigo21,"
        ZSql = ZSql + "Codigo22,"
        ZSql = ZSql + "Codigo23,"
        ZSql = ZSql + "Codigo24,"
        ZSql = ZSql + "Codigo25,"
        ZSql = ZSql + "Codigo26,"
        ZSql = ZSql + "Codigo27,"
        ZSql = ZSql + "Codigo28,"
        ZSql = ZSql + "Codigo29,"
        ZSql = ZSql + "Codigo30,"
        ZSql = ZSql + "Descri1,"
        ZSql = ZSql + "Descri2,"
        ZSql = ZSql + "Descri3,"
        ZSql = ZSql + "Descri4,"
        ZSql = ZSql + "Descri5,"
        ZSql = ZSql + "Descri6,"
        ZSql = ZSql + "Descri7,"
        ZSql = ZSql + "Descri8,"
        ZSql = ZSql + "Descri9,"
        ZSql = ZSql + "Descri10,"
        ZSql = ZSql + "Descri11,"
        ZSql = ZSql + "Descri12,"
        ZSql = ZSql + "Descri13,"
        ZSql = ZSql + "Descri14,"
        ZSql = ZSql + "Descri15,"
        ZSql = ZSql + "Descri16,"
        ZSql = ZSql + "Descri17,"
        ZSql = ZSql + "Descri18,"
        ZSql = ZSql + "Descri19,"
        ZSql = ZSql + "Descri20,"
        ZSql = ZSql + "Descri21,"
        ZSql = ZSql + "Descri22,"
        ZSql = ZSql + "Descri23,"
        ZSql = ZSql + "Descri24,"
        ZSql = ZSql + "Descri25,"
        ZSql = ZSql + "Descri26,"
        ZSql = ZSql + "Descri27,"
        ZSql = ZSql + "Descri28,"
        ZSql = ZSql + "Descri29,"
        ZSql = ZSql + "Descri30,"
        ZSql = ZSql + "Valor1,"
        ZSql = ZSql + "Valor2,"
        ZSql = ZSql + "Valor3,"
        ZSql = ZSql + "Valor4,"
        ZSql = ZSql + "Valor5,"
        ZSql = ZSql + "Valor6,"
        ZSql = ZSql + "Valor7,"
        ZSql = ZSql + "Valor8,"
        ZSql = ZSql + "Valor9,"
        ZSql = ZSql + "Valor10,"
        ZSql = ZSql + "ZValor1,"
        ZSql = ZSql + "ZValor2,"
        ZSql = ZSql + "ZValor3,"
        ZSql = ZSql + "ZValor4,"
        ZSql = ZSql + "ZValor5,"
        ZSql = ZSql + "ZValor6,"
        ZSql = ZSql + "ZValor7,"
        ZSql = ZSql + "ZValor8,"
        ZSql = ZSql + "ZValor9,"
        ZSql = ZSql + "ZValor10,"
        ZSql = ZSql + "ZValor121,"
        ZSql = ZSql + "ZValor122,"
        ZSql = ZSql + "ZValor123,"
        ZSql = ZSql + "ZValor124,"
        ZSql = ZSql + "ZValor125,"
        ZSql = ZSql + "ZValor126,"
        ZSql = ZSql + "ZValor127,"
        ZSql = ZSql + "ZValor128,"
        ZSql = ZSql + "ZValor129,"
        ZSql = ZSql + "ZValor130,"
        ZSql = ZSql + "Version ,"
        ZSql = ZSql + "Responsable,"
        ZSql = ZSql + "Fecha )"
        ZSql = ZSql + "Values ("
        ZSql = ZSql + "'" + ZCodigo + "',"
        ZSql = ZSql + "'" + ZDescripcion + "',"
        ZSql = ZSql + "'" + ZEnsayo(1) + "',"
        ZSql = ZSql + "'" + ZEnsayo(2) + "',"
        ZSql = ZSql + "'" + ZEnsayo(3) + "',"
        ZSql = ZSql + "'" + ZEnsayo(4) + "',"
        ZSql = ZSql + "'" + ZEnsayo(5) + "',"
        ZSql = ZSql + "'" + ZEnsayo(6) + "',"
        ZSql = ZSql + "'" + ZEnsayo(7) + "',"
        ZSql = ZSql + "'" + ZEnsayo(8) + "',"
        ZSql = ZSql + "'" + ZEnsayo(9) + "',"
        ZSql = ZSql + "'" + ZEnsayo(10) + "',"
        ZSql = ZSql + "'" + ZEnsayo(11) + "',"
        ZSql = ZSql + "'" + ZEnsayo(12) + "',"
        ZSql = ZSql + "'" + ZEnsayo(13) + "',"
        ZSql = ZSql + "'" + ZEnsayo(14) + "',"
        ZSql = ZSql + "'" + ZEnsayo(15) + "',"
        ZSql = ZSql + "'" + ZEnsayo(16) + "',"
        ZSql = ZSql + "'" + ZEnsayo(17) + "',"
        ZSql = ZSql + "'" + ZEnsayo(18) + "',"
        ZSql = ZSql + "'" + ZEnsayo(19) + "',"
        ZSql = ZSql + "'" + ZEnsayo(20) + "',"
        ZSql = ZSql + "'" + ZEnsayo(21) + "',"
        ZSql = ZSql + "'" + ZEnsayo(22) + "',"
        ZSql = ZSql + "'" + ZEnsayo(23) + "',"
        ZSql = ZSql + "'" + ZEnsayo(24) + "',"
        ZSql = ZSql + "'" + ZEnsayo(25) + "',"
        ZSql = ZSql + "'" + ZEnsayo(26) + "',"
        ZSql = ZSql + "'" + ZEnsayo(27) + "',"
        ZSql = ZSql + "'" + ZEnsayo(28) + "',"
        ZSql = ZSql + "'" + ZEnsayo(29) + "',"
        ZSql = ZSql + "'" + ZEnsayo(30) + "',"
        ZSql = ZSql + "'" + ZDescri(1) + "',"
        ZSql = ZSql + "'" + ZDescri(2) + "',"
        ZSql = ZSql + "'" + ZDescri(3) + "',"
        ZSql = ZSql + "'" + ZDescri(4) + "',"
        ZSql = ZSql + "'" + ZDescri(5) + "',"
        ZSql = ZSql + "'" + ZDescri(6) + "',"
        ZSql = ZSql + "'" + ZDescri(7) + "',"
        ZSql = ZSql + "'" + ZDescri(8) + "',"
        ZSql = ZSql + "'" + ZDescri(9) + "',"
        ZSql = ZSql + "'" + ZDescri(10) + "',"
        ZSql = ZSql + "'" + ZDescri(11) + "',"
        ZSql = ZSql + "'" + ZDescri(12) + "',"
        ZSql = ZSql + "'" + ZDescri(13) + "',"
        ZSql = ZSql + "'" + ZDescri(14) + "',"
        ZSql = ZSql + "'" + ZDescri(15) + "',"
        ZSql = ZSql + "'" + ZDescri(16) + "',"
        ZSql = ZSql + "'" + ZDescri(17) + "',"
        ZSql = ZSql + "'" + ZDescri(18) + "',"
        ZSql = ZSql + "'" + ZDescri(19) + "',"
        ZSql = ZSql + "'" + ZDescri(20) + "',"
        ZSql = ZSql + "'" + ZDescri(21) + "',"
        ZSql = ZSql + "'" + ZDescri(22) + "',"
        ZSql = ZSql + "'" + ZDescri(23) + "',"
        ZSql = ZSql + "'" + ZDescri(24) + "',"
        ZSql = ZSql + "'" + ZDescri(25) + "',"
        ZSql = ZSql + "'" + ZDescri(26) + "',"
        ZSql = ZSql + "'" + ZDescri(27) + "',"
        ZSql = ZSql + "'" + ZDescri(28) + "',"
        ZSql = ZSql + "'" + ZDescri(29) + "',"
        ZSql = ZSql + "'" + ZDescri(30) + "',"
        ZSql = ZSql + "'" + ZValor(1) + "',"
        ZSql = ZSql + "'" + ZValor(2) + "',"
        ZSql = ZSql + "'" + ZValor(3) + "',"
        ZSql = ZSql + "'" + ZValor(4) + "',"
        ZSql = ZSql + "'" + ZValor(5) + "',"
        ZSql = ZSql + "'" + ZValor(6) + "',"
        ZSql = ZSql + "'" + ZValor(7) + "',"
        ZSql = ZSql + "'" + ZValor(8) + "',"
        ZSql = ZSql + "'" + ZValor(9) + "',"
        ZSql = ZSql + "'" + ZValor(10) + "',"
        ZSql = ZSql + "'" + ZValor(11) + "',"
        ZSql = ZSql + "'" + ZValor(12) + "',"
        ZSql = ZSql + "'" + ZValor(13) + "',"
        ZSql = ZSql + "'" + ZValor(14) + "',"
        ZSql = ZSql + "'" + ZValor(15) + "',"
        ZSql = ZSql + "'" + ZValor(16) + "',"
        ZSql = ZSql + "'" + ZValor(17) + "',"
        ZSql = ZSql + "'" + ZValor(18) + "',"
        ZSql = ZSql + "'" + ZValor(19) + "',"
        ZSql = ZSql + "'" + ZValor(20) + "',"
        ZSql = ZSql + "'" + ZValor(21) + "',"
        ZSql = ZSql + "'" + ZValor(22) + "',"
        ZSql = ZSql + "'" + ZValor(23) + "',"
        ZSql = ZSql + "'" + ZValor(24) + "',"
        ZSql = ZSql + "'" + ZValor(25) + "',"
        ZSql = ZSql + "'" + ZValor(26) + "',"
        ZSql = ZSql + "'" + ZValor(27) + "',"
        ZSql = ZSql + "'" + ZValor(28) + "',"
        ZSql = ZSql + "'" + ZValor(29) + "',"
        ZSql = ZSql + "'" + ZValor(30) + "',"
        ZSql = ZSql + "'" + ZZVersion + "',"
        ZSql = ZSql + "'" + ZZDesOperador + "',"
        ZSql = ZSql + "'" + ZZFecha + "')"
        
        spListaEspe = ZSql
        Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
        
        
    Next Ciclo
    
    ZSql = ""
    ZSql = ZSql + "UPDATE ListaEspe SET "
    ZSql = ZSql + "ControlCambio = " + "'" + ControlCambio.Text + "'"
    spListaEspe = ZSql
    Set rstListaEspe = db.OpenRecordset(spListaEspe, dbOpenSnapshot, dbSQLPassThrough)
    
    lista.WindowTitle = "Listado de Especificaciones de Materia Prima (Unificado)"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height

    lista.Destination = 1
    
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)
    
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            lista.ReportFileName = "ListaEspe.rpt"
        Case Else
            lista.ReportFileName = "ListaEspePelli.rpt"
    End Select
    
    Rem Lista.SQLQuery = "SELECT ListaEspe.Codigo, ListaEspe.Descripcion, " _
    rem             + "ListaEspe.Codigo1, ListaEspe.Codigo2, ListaEspe.Codigo3, ListaEspe.Codigo4, ListaEspe.Codigo5, ListaEspe.Codigo6, ListaEspe.Codigo7, ListaEspe.Codigo8, ListaEspe.Codigo9, ListaEspe.Codigo10, " _
    rem             + "ListaEspe.Codigo11, ListaEspe.Codigo12, ListaEspe.Codigo13, ListaEspe.Codigo14, ListaEspe.Codigo15, ListaEspe.Codigo16, ListaEspe.Codigo17, ListaEspe.Codigo18, ListaEspe.Codigo19, ListaEspe.Codigo20, " _
    rem             + "ListaEspe.Descri1, ListaEspe.Descri2, ListaEspe.Descri3, ListaEspe.Descri4, ListaEspe.Descri5, ListaEspe.Descri6, ListaEspe.Descri7, ListaEspe.Descri8, ListaEspe.Descri9, ListaEspe.Descri10, " _
    rem             + "ListaEspe.Descri11, ListaEspe.Descri12, ListaEspe.Descri13, ListaEspe.Descri14, ListaEspe.Descri15, ListaEspe.Descri16, ListaEspe.Descri17, ListaEspe.Descri18, ListaEspe.Descri19, ListaEspe.Descri20, " _
    rem             + "ListaEspe.Valor1, ListaEspe.Valor2, ListaEspe.Valor3, ListaEspe.Valor4, ListaEspe.Valor5, ListaEspe.Valor6, ListaEspe.Valor7, ListaEspe.Valor8, ListaEspe.Valor9, ListaEspe.Valor10, " _
    rem             + "ListaEspe.ZValor1, ListaEspe.ZValor2, ListaEspe.ZValor3, ListaEspe.ZValor4, ListaEspe.ZValor5, ListaEspe.ZValor6, ListaEspe.ZValor7, ListaEspe.ZValor8, ListaEspe.ZValor9, ListaEspe.ZValor10, " _
    rem             + "ListaEspe.Version, ListaEspe.Responsable, ListaEspe.Fecha, ListaEspe.ControlCambio " _
    rem             + "From " _
    rem             + DSQ + ".dbo.ListaEspe ListaEspe " _
    rem             + "Where " _
    rem             + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
    rem             + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    
    lista.SQLQuery = "SELECT * " _
                + "From " _
                + DSQ + ".dbo.ListaEspe ListaEspe " _
                + "Where " _
                + "ListaEspe.Codigo >= '" + Desde.Text + "' AND " _
                + "ListaEspe.Codigo <= '" + Hasta.Text + "'"
    
    lista.Connect = Connect()
    
    lista.Action = 1
    Frame2.Visible = False
    
    Call Conecta_Empresa
    
End Sub

Private Sub Cancela_click()
    Frame2.Visible = False
End Sub

Private Sub cmdAdd_Click()

    If Trim(ControlCambio.Text) = "" Then
        m$ = "Se debe informar el campo Control de Cambio"
        A% = MsgBox(m$, 0, "Especificaciones de Producto Terminado")
        Exit Sub
    End If

    If WGraba <> "S" Then
    
        ZZProceso = 0
        Call Ingresa_clave

               Else

        If Codigo.Text <> "" Then
    
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
        
            WProducto = Codigo.Text
            WEnsayo1 = Ensayo1.Text
            WEnsayo2 = Ensayo2.Text
            WEnsayo3 = Ensayo3.Text
            WEnsayo4 = Ensayo4.Text
            WEnsayo5 = Ensayo5.Text
            WEnsayo6 = Ensayo6.Text
            WEnsayo7 = Ensayo7.Text
            WEnsayo8 = Ensayo8.Text
            WEnsayo9 = Ensayo9.Text
            WEnsayo10 = Ensayo10.Text
            WEnsayo11 = Ensayo11.Text
            WEnsayo12 = Ensayo12.Text
            WEnsayo13 = Ensayo13.Text
            WEnsayo14 = Ensayo14.Text
            WEnsayo15 = Ensayo15.Text
            WEnsayo16 = Ensayo16.Text
            WEnsayo17 = Ensayo17.Text
            WEnsayo18 = Ensayo18.Text
            WEnsayo19 = Ensayo19.Text
            WEnsayo20 = Ensayo20.Text
            WEnsayo21 = Ensayo21.Text
            WEnsayo22 = Ensayo22.Text
            WEnsayo23 = Ensayo23.Text
            WEnsayo24 = Ensayo24.Text
            WEnsayo25 = Ensayo25.Text
            WEnsayo26 = Ensayo26.Text
            WEnsayo27 = Ensayo27.Text
            WEnsayo28 = Ensayo28.Text
            WEnsayo29 = Ensayo29.Text
            WEnsayo30 = Ensayo30.Text
            
            WValor1 = Valor1.Text
            WValor2 = valor2.Text
            WValor3 = Valor3.Text
            WValor4 = valor4.Text
            WValor5 = valor5.Text
            WValor6 = valor6.Text
            WValor7 = valor7.Text
            WValor8 = valor8.Text
            WValor9 = valor9.Text
            WValor10 = valor10.Text
            WValor11 = Valor11.Text
            WValor12 = Valor12.Text
            WValor13 = Valor13.Text
            WValor14 = Valor14.Text
            WValor15 = Valor15.Text
            WValor16 = Valor16.Text
            WValor17 = Valor17.Text
            WValor18 = Valor18.Text
            WValor19 = Valor19.Text
            WValor20 = Valor20.Text
            WValor21 = Valor21.Text
            WValor22 = Valor22.Text
            WValor23 = Valor23.Text
            WValor24 = Valor24.Text
            WValor25 = Valor25.Text
            WValor26 = Valor26.Text
            WValor27 = Valor27.Text
            WValor28 = Valor28.Text
            WValor29 = Valor29.Text
            WValor30 = Valor30.Text
            
            WDesde1 = Desde1.Text
            WDesde2 = Desde2.Text
            WDesde3 = Desde3.Text
            WDesde4 = Desde4.Text
            WDesde5 = Desde5.Text
            WDesde6 = Desde6.Text
            WDesde7 = Desde7.Text
            WDesde8 = Desde8.Text
            WDesde9 = Desde9.Text
            WDesde10 = Desde10.Text
            WDesde11 = Desde11.Text
            WDesde12 = Desde12.Text
            WDesde13 = Desde13.Text
            WDesde14 = Desde14.Text
            WDesde15 = Desde15.Text
            WDesde16 = Desde16.Text
            WDesde17 = Desde17.Text
            WDesde18 = Desde18.Text
            WDesde19 = Desde19.Text
            WDesde20 = Desde20.Text
            WDesde21 = Desde21.Text
            WDesde22 = Desde22.Text
            WDesde23 = Desde23.Text
            WDesde24 = Desde24.Text
            WDesde25 = Desde25.Text
            WDesde26 = Desde26.Text
            WDesde27 = Desde27.Text
            WDesde28 = Desde28.Text
            WDesde29 = Desde29.Text
            WDesde30 = Desde30.Text
            
            WHasta1 = Hasta1.Text
            WHasta2 = Hasta2.Text
            WHasta3 = Hasta3.Text
            WHasta4 = Hasta4.Text
            WHasta5 = Hasta5.Text
            WHasta6 = Hasta6.Text
            WHasta7 = Hasta7.Text
            WHasta8 = Hasta8.Text
            WHasta9 = Hasta9.Text
            WHasta10 = Hasta10.Text
            WHasta11 = Hasta11.Text
            WHasta12 = Hasta12.Text
            WHasta13 = Hasta13.Text
            WHasta14 = Hasta14.Text
            WHasta15 = Hasta15.Text
            WHasta16 = Hasta16.Text
            WHasta17 = Hasta17.Text
            WHasta18 = Hasta18.Text
            WHasta19 = Hasta19.Text
            WHasta20 = Hasta20.Text
            WHasta21 = Hasta21.Text
            WHasta22 = Hasta22.Text
            WHasta23 = Hasta23.Text
            WHasta24 = Hasta24.Text
            WHasta25 = Hasta25.Text
            WHasta26 = Hasta26.Text
            WHasta27 = Hasta27.Text
            WHasta28 = Hasta28.Text
            WHasta29 = Hasta29.Text
            WHasta30 = Hasta30.Text
            
            WDate = Date$
            
            Sql1 = "Select *"
            Sql2 = " FROM EspecificacionesUnifica"
            Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
            spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnifica.RecordCount > 0 Then
        
                ZProducto = rstEspecificacionesUnifica!Producto
                
                ZEnsayo1 = Str$(rstEspecificacionesUnifica!Ensayo1)
                ZEnsayo2 = Str$(rstEspecificacionesUnifica!Ensayo2)
                ZEnsayo3 = Str$(rstEspecificacionesUnifica!Ensayo3)
                ZEnsayo4 = Str$(rstEspecificacionesUnifica!Ensayo4)
                ZEnsayo5 = Str$(rstEspecificacionesUnifica!Ensayo5)
                ZEnsayo6 = Str$(rstEspecificacionesUnifica!Ensayo6)
                ZEnsayo7 = Str$(rstEspecificacionesUnifica!Ensayo7)
                ZEnsayo8 = Str$(rstEspecificacionesUnifica!Ensayo8)
                ZEnsayo9 = Str$(rstEspecificacionesUnifica!Ensayo9)
                ZEnsayo10 = Str$(rstEspecificacionesUnifica!Ensayo10)
                ZEnsayo11 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
                ZEnsayo12 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
                ZEnsayo13 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
                ZEnsayo14 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
                ZEnsayo15 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
                ZEnsayo16 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
                ZEnsayo17 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
                ZEnsayo18 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
                ZEnsayo19 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
                ZEnsayo20 = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
                
                ZValor1 = rstEspecificacionesUnifica!Valor1
                ZValor2 = rstEspecificacionesUnifica!valor2
                ZValor3 = rstEspecificacionesUnifica!Valor3
                ZValor4 = rstEspecificacionesUnifica!valor4
                ZValor5 = rstEspecificacionesUnifica!valor5
                ZValor6 = rstEspecificacionesUnifica!valor6
                ZValor7 = rstEspecificacionesUnifica!valor7
                ZValor8 = rstEspecificacionesUnifica!valor8
                ZValor9 = rstEspecificacionesUnifica!valor9
                ZValor10 = rstEspecificacionesUnifica!valor10
                ZValor11 = IIf(IsNull(rstEspecificacionesUnifica!Valor11), "", rstEspecificacionesUnifica!Valor11)
                ZValor12 = IIf(IsNull(rstEspecificacionesUnifica!Valor12), "", rstEspecificacionesUnifica!Valor12)
                ZValor13 = IIf(IsNull(rstEspecificacionesUnifica!Valor13), "", rstEspecificacionesUnifica!Valor13)
                ZValor14 = IIf(IsNull(rstEspecificacionesUnifica!Valor14), "", rstEspecificacionesUnifica!Valor14)
                ZValor15 = IIf(IsNull(rstEspecificacionesUnifica!Valor15), "", rstEspecificacionesUnifica!Valor15)
                ZValor16 = IIf(IsNull(rstEspecificacionesUnifica!Valor16), "", rstEspecificacionesUnifica!Valor16)
                ZValor17 = IIf(IsNull(rstEspecificacionesUnifica!Valor17), "", rstEspecificacionesUnifica!Valor17)
                ZValor18 = IIf(IsNull(rstEspecificacionesUnifica!Valor18), "", rstEspecificacionesUnifica!Valor18)
                ZValor19 = IIf(IsNull(rstEspecificacionesUnifica!Valor19), "", rstEspecificacionesUnifica!Valor19)
                ZValor20 = IIf(IsNull(rstEspecificacionesUnifica!Valor20), "", rstEspecificacionesUnifica!Valor20)
                
                ZDesde1 = IIf(IsNull(rstEspecificacionesUnifica!Desde1), "", rstEspecificacionesUnifica!Desde1)
                ZDesde2 = IIf(IsNull(rstEspecificacionesUnifica!Desde2), "", rstEspecificacionesUnifica!Desde2)
                ZDesde3 = IIf(IsNull(rstEspecificacionesUnifica!Desde3), "", rstEspecificacionesUnifica!Desde3)
                ZDesde4 = IIf(IsNull(rstEspecificacionesUnifica!Desde4), "", rstEspecificacionesUnifica!Desde4)
                ZDesde5 = IIf(IsNull(rstEspecificacionesUnifica!Desde5), "", rstEspecificacionesUnifica!Desde5)
                ZDesde6 = IIf(IsNull(rstEspecificacionesUnifica!Desde6), "", rstEspecificacionesUnifica!Desde6)
                ZDesde7 = IIf(IsNull(rstEspecificacionesUnifica!Desde7), "", rstEspecificacionesUnifica!Desde7)
                ZDesde8 = IIf(IsNull(rstEspecificacionesUnifica!Desde8), "", rstEspecificacionesUnifica!Desde8)
                ZDesde9 = IIf(IsNull(rstEspecificacionesUnifica!Desde9), "", rstEspecificacionesUnifica!Desde9)
                ZDesde10 = IIf(IsNull(rstEspecificacionesUnifica!Desde10), "", rstEspecificacionesUnifica!Desde10)
                ZDesde11 = IIf(IsNull(rstEspecificacionesUnifica!Desde11), "", rstEspecificacionesUnifica!Desde11)
                ZDesde12 = IIf(IsNull(rstEspecificacionesUnifica!Desde12), "", rstEspecificacionesUnifica!Desde12)
                ZDesde13 = IIf(IsNull(rstEspecificacionesUnifica!Desde13), "", rstEspecificacionesUnifica!Desde13)
                ZDesde14 = IIf(IsNull(rstEspecificacionesUnifica!Desde14), "", rstEspecificacionesUnifica!Desde14)
                ZDesde15 = IIf(IsNull(rstEspecificacionesUnifica!Desde15), "", rstEspecificacionesUnifica!Desde15)
                ZDesde16 = IIf(IsNull(rstEspecificacionesUnifica!Desde16), "", rstEspecificacionesUnifica!Desde16)
                ZDesde17 = IIf(IsNull(rstEspecificacionesUnifica!Desde17), "", rstEspecificacionesUnifica!Desde17)
                ZDesde18 = IIf(IsNull(rstEspecificacionesUnifica!Desde18), "", rstEspecificacionesUnifica!Desde18)
                ZDesde19 = IIf(IsNull(rstEspecificacionesUnifica!Desde19), "", rstEspecificacionesUnifica!Desde19)
                ZDesde20 = IIf(IsNull(rstEspecificacionesUnifica!Desde20), "", rstEspecificacionesUnifica!Desde20)
                
                ZHasta1 = IIf(IsNull(rstEspecificacionesUnifica!Hasta1), "", rstEspecificacionesUnifica!Hasta1)
                ZHasta2 = IIf(IsNull(rstEspecificacionesUnifica!Hasta2), "", rstEspecificacionesUnifica!Hasta2)
                ZHasta3 = IIf(IsNull(rstEspecificacionesUnifica!Hasta3), "", rstEspecificacionesUnifica!Hasta3)
                ZHasta4 = IIf(IsNull(rstEspecificacionesUnifica!Hasta4), "", rstEspecificacionesUnifica!Hasta4)
                ZHasta5 = IIf(IsNull(rstEspecificacionesUnifica!Hasta5), "", rstEspecificacionesUnifica!Hasta5)
                ZHasta6 = IIf(IsNull(rstEspecificacionesUnifica!Hasta6), "", rstEspecificacionesUnifica!Hasta6)
                ZHasta7 = IIf(IsNull(rstEspecificacionesUnifica!Hasta7), "", rstEspecificacionesUnifica!Hasta7)
                ZHasta8 = IIf(IsNull(rstEspecificacionesUnifica!Hasta8), "", rstEspecificacionesUnifica!Hasta8)
                ZHasta9 = IIf(IsNull(rstEspecificacionesUnifica!Hasta9), "", rstEspecificacionesUnifica!Hasta9)
                ZHasta10 = IIf(IsNull(rstEspecificacionesUnifica!Hasta10), "", rstEspecificacionesUnifica!Hasta10)
                ZHasta11 = IIf(IsNull(rstEspecificacionesUnifica!Hasta11), "", rstEspecificacionesUnifica!Hasta11)
                ZHasta12 = IIf(IsNull(rstEspecificacionesUnifica!Hasta12), "", rstEspecificacionesUnifica!Hasta12)
                ZHasta13 = IIf(IsNull(rstEspecificacionesUnifica!Hasta13), "", rstEspecificacionesUnifica!Hasta13)
                ZHasta14 = IIf(IsNull(rstEspecificacionesUnifica!Hasta14), "", rstEspecificacionesUnifica!Hasta14)
                ZHasta15 = IIf(IsNull(rstEspecificacionesUnifica!Hasta15), "", rstEspecificacionesUnifica!Hasta15)
                ZHasta16 = IIf(IsNull(rstEspecificacionesUnifica!Hasta16), "", rstEspecificacionesUnifica!Hasta16)
                ZHasta17 = IIf(IsNull(rstEspecificacionesUnifica!Hasta17), "", rstEspecificacionesUnifica!Hasta17)
                ZHasta18 = IIf(IsNull(rstEspecificacionesUnifica!Hasta18), "", rstEspecificacionesUnifica!Hasta18)
                ZHasta19 = IIf(IsNull(rstEspecificacionesUnifica!Hasta19), "", rstEspecificacionesUnifica!Hasta19)
                ZHasta20 = IIf(IsNull(rstEspecificacionesUnifica!Hasta20), "", rstEspecificacionesUnifica!Hasta20)
                
                ZVersion = Str$(rstEspecificacionesUnifica!Version)
                ZFechaInicio = rstEspecificacionesUnifica!fecha
                ZFechaFinal = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
        
                rstEspecificacionesUnifica.Close
            
                ZEnsayo21 = ""
                ZEnsayo22 = ""
                ZEnsayo23 = ""
                ZEnsayo24 = ""
                ZEnsayo25 = ""
                ZEnsayo26 = ""
                ZEnsayo27 = ""
                ZEnsayo28 = ""
                ZEnsayo29 = ""
                ZEnsayo30 = ""
                
                ZValor21 = ""
                ZValor22 = ""
                ZValor23 = ""
                ZValor24 = ""
                ZValor25 = ""
                ZValor26 = ""
                ZValor27 = ""
                ZValor28 = ""
                ZValor29 = ""
                ZValor30 = ""
                
                ZDesde21 = ""
                ZDesde22 = ""
                ZDesde23 = ""
                ZDesde24 = ""
                ZDesde25 = ""
                ZDesde26 = ""
                ZDesde27 = ""
                ZDesde28 = ""
                ZDesde29 = ""
                ZDesde30 = ""
                
                ZHasta21 = ""
                ZHasta22 = ""
                ZHasta23 = ""
                ZHasta24 = ""
                ZHasta25 = ""
                ZHasta26 = ""
                ZHasta27 = ""
                ZHasta28 = ""
                ZHasta29 = ""
                ZHasta30 = ""
                
                Sql1 = "Select *"
                Sql2 = " FROM EspecificacionesUnificaIII"
                Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Codigo.Text + "'"
                spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
                Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
                If rstEspecificacionesUnificaIII.RecordCount > 0 Then
                
                    ZEnsayo21 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
                    ZEnsayo22 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
                    ZEnsayo23 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
                    ZEnsayo24 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
                    ZEnsayo25 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
                    ZEnsayo26 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
                    ZEnsayo27 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
                    ZEnsayo28 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
                    ZEnsayo29 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
                    ZEnsayo30 = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
                    
                    ZValor21 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor21), "", rstEspecificacionesUnificaIII!Valor21)
                    ZValor22 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor22), "", rstEspecificacionesUnificaIII!Valor22)
                    ZValor23 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor23), "", rstEspecificacionesUnificaIII!Valor23)
                    ZValor24 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor24), "", rstEspecificacionesUnificaIII!Valor24)
                    ZValor25 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor25), "", rstEspecificacionesUnificaIII!Valor25)
                    ZValor26 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor26), "", rstEspecificacionesUnificaIII!Valor26)
                    ZValor27 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor27), "", rstEspecificacionesUnificaIII!Valor27)
                    ZValor28 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor28), "", rstEspecificacionesUnificaIII!Valor28)
                    ZValor29 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor29), "", rstEspecificacionesUnificaIII!Valor29)
                    ZValor30 = IIf(IsNull(rstEspecificacionesUnificaIII!Valor30), "", rstEspecificacionesUnificaIII!Valor30)
                    
                    ZDesde21 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde21), "", rstEspecificacionesUnificaIII!Desde21)
                    ZDesde22 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde22), "", rstEspecificacionesUnificaIII!Desde22)
                    ZDesde23 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde23), "", rstEspecificacionesUnificaIII!Desde23)
                    ZDesde24 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde24), "", rstEspecificacionesUnificaIII!Desde24)
                    ZDesde25 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde25), "", rstEspecificacionesUnificaIII!Desde25)
                    ZDesde26 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde26), "", rstEspecificacionesUnificaIII!Desde26)
                    ZDesde27 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde27), "", rstEspecificacionesUnificaIII!Desde27)
                    ZDesde28 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde28), "", rstEspecificacionesUnificaIII!Desde28)
                    ZDesde29 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde29), "", rstEspecificacionesUnificaIII!Desde29)
                    ZDesde30 = IIf(IsNull(rstEspecificacionesUnificaIII!Desde30), "", rstEspecificacionesUnificaIII!Desde30)
                   
                    ZHasta21 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta21), "", rstEspecificacionesUnificaIII!Hasta21)
                    ZHasta22 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta22), "", rstEspecificacionesUnificaIII!Hasta22)
                    ZHasta23 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta23), "", rstEspecificacionesUnificaIII!Hasta23)
                    ZHasta24 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta24), "", rstEspecificacionesUnificaIII!Hasta24)
                    ZHasta25 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta25), "", rstEspecificacionesUnificaIII!Hasta25)
                    ZHasta26 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta26), "", rstEspecificacionesUnificaIII!Hasta26)
                    ZHasta27 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta27), "", rstEspecificacionesUnificaIII!Hasta27)
                    ZHasta28 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta28), "", rstEspecificacionesUnificaIII!Hasta28)
                    ZHasta29 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta29), "", rstEspecificacionesUnificaIII!Hasta29)
                    ZHasta30 = IIf(IsNull(rstEspecificacionesUnificaIII!Hasta30), "", rstEspecificacionesUnificaIII!Hasta30)
                
                    rstEspecificacionesUnificaIII.Close
                    
                End If
                
                ZValor1 = Trim(ZValor1)
                ZValor2 = Trim(ZValor2)
                ZValor3 = Trim(ZValor3)
                ZValor4 = Trim(ZValor4)
                ZValor5 = Trim(ZValor5)
                ZValor6 = Trim(ZValor6)
                ZValor7 = Trim(ZValor7)
                ZValor8 = Trim(ZValor8)
                ZValor9 = Trim(ZValor9)
                ZValor10 = Trim(ZValor10)
                ZValor11 = Trim(ZValor11)
                ZValor12 = Trim(ZValor12)
                ZValor13 = Trim(ZValor13)
                ZValor14 = Trim(ZValor14)
                ZValor15 = Trim(ZValor15)
                ZValor16 = Trim(ZValor16)
                ZValor17 = Trim(ZValor17)
                ZValor18 = Trim(ZValor18)
                ZValor19 = Trim(ZValor19)
                ZValor20 = Trim(ZValor20)
                ZValor21 = Trim(ZValor21)
                ZValor22 = Trim(ZValor22)
                ZValor23 = Trim(ZValor23)
                ZValor24 = Trim(ZValor24)
                ZValor25 = Trim(ZValor25)
                ZValor26 = Trim(ZValor26)
                ZValor27 = Trim(ZValor27)
                ZValor28 = Trim(ZValor28)
                ZValor29 = Trim(ZValor29)
                ZValor30 = Trim(ZValor30)
                
                ZValor1 = Left$(ZValor1, 50)
                ZValor2 = Left$(ZValor2, 50)
                ZValor3 = Left$(ZValor3, 50)
                ZValor4 = Left$(ZValor4, 50)
                ZValor5 = Left$(ZValor5, 50)
                ZValor6 = Left$(ZValor6, 50)
                ZValor7 = Left$(ZValor7, 50)
                ZValor8 = Left$(ZValor8, 50)
                ZValor9 = Left$(ZValor9, 50)
                ZValor10 = Left$(ZValor10, 50)
                ZValor11 = Left$(ZValor11, 50)
                ZValor12 = Left$(ZValor12, 50)
                ZValor13 = Left$(ZValor13, 50)
                ZValor14 = Left$(ZValor14, 50)
                ZValor15 = Left$(ZValor15, 50)
                ZValor16 = Left$(ZValor16, 50)
                ZValor17 = Left$(ZValor17, 50)
                ZValor18 = Left$(ZValor18, 50)
                ZValor19 = Left$(ZValor19, 50)
                ZValor20 = Left$(ZValor20, 50)
                ZValor21 = Left$(ZValor21, 50)
                ZValor22 = Left$(ZValor22, 50)
                ZValor23 = Left$(ZValor23, 50)
                ZValor24 = Left$(ZValor24, 50)
                ZValor25 = Left$(ZValor25, 50)
                ZValor26 = Left$(ZValor26, 50)
                ZValor27 = Left$(ZValor27, 50)
                ZValor28 = Left$(ZValor28, 50)
                ZValor29 = Left$(ZValor29, 50)
                ZValor30 = Left$(ZValor30, 50)
                
                
                
                
                
                Rem
                Rem graba los datos para la version
                Rem
            
                Call Ceros(ZVersion, 4)
                ZClave = ZVersion + ZProducto
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EspecificacionesUnificaVersion ("
                ZSql = ZSql + "Clave, "
                ZSql = ZSql + "Version, "
                ZSql = ZSql + "Producto, "
                ZSql = ZSql + "Ensayo1, Valor1, "
                ZSql = ZSql + "Ensayo2, Valor2, "
                ZSql = ZSql + "Ensayo3, Valor3, "
                ZSql = ZSql + "Ensayo4, Valor4, "
                ZSql = ZSql + "Ensayo5, Valor5, "
                ZSql = ZSql + "Ensayo6, Valor6, "
                ZSql = ZSql + "Ensayo7, Valor7, "
                ZSql = ZSql + "Ensayo8, Valor8, "
                ZSql = ZSql + "Ensayo9, Valor9, "
                ZSql = ZSql + "Ensayo10, Valor10, "
                ZSql = ZSql + "Ensayo11, ZValor1, "
                ZSql = ZSql + "Ensayo12, ZValor2, "
                ZSql = ZSql + "Ensayo13, ZValor3, "
                ZSql = ZSql + "Ensayo14, ZValor4, "
                ZSql = ZSql + "Ensayo15, ZValor5, "
                ZSql = ZSql + "Ensayo16, ZValor6, "
                ZSql = ZSql + "Ensayo17, ZValor7, "
                ZSql = ZSql + "Ensayo18, ZValor8, "
                ZSql = ZSql + "Ensayo19, ZValor9, "
                ZSql = ZSql + "Ensayo20, ZValor10, "
                ZSql = ZSql + "ControlCambio, "
                ZSql = ZSql + "FechaInicio, "
                ZSql = ZSql + "FechaFinal) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZClave + "',"
                ZSql = ZSql + "'" + ZVersion + "',"
                ZSql = ZSql + "'" + ZProducto + "',"
                ZSql = ZSql + "'" + ZEnsayo1 + "'," + "'" + ZValor1 + "',"
                ZSql = ZSql + "'" + ZEnsayo2 + "'," + "'" + ZValor2 + "',"
                ZSql = ZSql + "'" + ZEnsayo3 + "'," + "'" + ZValor3 + "',"
                ZSql = ZSql + "'" + ZEnsayo4 + "'," + "'" + ZValor4 + "',"
                ZSql = ZSql + "'" + ZEnsayo5 + "'," + "'" + ZValor5 + "',"
                ZSql = ZSql + "'" + ZEnsayo6 + "'," + "'" + ZValor6 + "',"
                ZSql = ZSql + "'" + ZEnsayo7 + "'," + "'" + ZValor7 + "',"
                ZSql = ZSql + "'" + ZEnsayo8 + "'," + "'" + ZValor8 + "',"
                ZSql = ZSql + "'" + ZEnsayo9 + "'," + "'" + ZValor9 + "',"
                ZSql = ZSql + "'" + ZEnsayo10 + "'," + "'" + ZValor10 + "',"
                ZSql = ZSql + "'" + ZEnsayo11 + "'," + "'" + ZValor11 + "',"
                ZSql = ZSql + "'" + ZEnsayo12 + "'," + "'" + ZValor12 + "',"
                ZSql = ZSql + "'" + ZEnsayo13 + "'," + "'" + ZValor13 + "',"
                ZSql = ZSql + "'" + ZEnsayo14 + "'," + "'" + ZValor14 + "',"
                ZSql = ZSql + "'" + ZEnsayo15 + "'," + "'" + ZValor15 + "',"
                ZSql = ZSql + "'" + ZEnsayo16 + "'," + "'" + ZValor16 + "',"
                ZSql = ZSql + "'" + ZEnsayo17 + "'," + "'" + ZValor17 + "',"
                ZSql = ZSql + "'" + ZEnsayo18 + "'," + "'" + ZValor18 + "',"
                ZSql = ZSql + "'" + ZEnsayo19 + "'," + "'" + ZValor19 + "',"
                ZSql = ZSql + "'" + ZEnsayo20 + "'," + "'" + ZValor20 + "',"
                ZSql = ZSql + "'" + ControlCambio.Text + "',"
                ZSql = ZSql + "'" + ZFechaInicio + "',"
                ZSql = ZSql + "'" + ZFechaFinal + "')"
            
                spEspecificacionesUnificaVersion = ZSql
                Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecificacionesUnificaVersion SET "
                ZSql = ZSql + "Desde1 = " + "'" + ZDesde1 + "',"
                ZSql = ZSql + "Hasta1 = " + "'" + ZHasta1 + "',"
                ZSql = ZSql + "Desde2 = " + "'" + ZDesde2 + "',"
                ZSql = ZSql + "Hasta2 = " + "'" + ZHasta2 + "',"
                ZSql = ZSql + "Desde3 = " + "'" + ZDesde3 + "',"
                ZSql = ZSql + "Hasta3 = " + "'" + ZHasta3 + "',"
                ZSql = ZSql + "Desde4 = " + "'" + ZDesde4 + "',"
                ZSql = ZSql + "Hasta4 = " + "'" + ZHasta4 + "',"
                ZSql = ZSql + "Desde5 = " + "'" + ZDesde5 + "',"
                ZSql = ZSql + "Hasta5 = " + "'" + ZHasta5 + "',"
                ZSql = ZSql + "Desde6 = " + "'" + ZDesde6 + "',"
                ZSql = ZSql + "Hasta6 = " + "'" + ZHasta6 + "',"
                ZSql = ZSql + "Desde7 = " + "'" + ZDesde7 + "',"
                ZSql = ZSql + "Hasta7 = " + "'" + ZHasta7 + "',"
                ZSql = ZSql + "Desde8 = " + "'" + ZDesde8 + "',"
                ZSql = ZSql + "Hasta8 = " + "'" + ZHasta8 + "',"
                ZSql = ZSql + "Desde9 = " + "'" + ZDesde9 + "',"
                ZSql = ZSql + "Hasta9 = " + "'" + ZHasta9 + "',"
                ZSql = ZSql + "Desde10 = " + "'" + ZDesde10 + "',"
                ZSql = ZSql + "Hasta10 = " + "'" + ZHasta10 + "',"
                ZSql = ZSql + "Desde11 = " + "'" + ZDesde11 + "',"
                ZSql = ZSql + "Hasta11 = " + "'" + ZHasta11 + "',"
                ZSql = ZSql + "Desde12 = " + "'" + ZDesde12 + "',"
                ZSql = ZSql + "Hasta12 = " + "'" + ZHasta12 + "',"
                ZSql = ZSql + "Desde13 = " + "'" + ZDesde13 + "',"
                ZSql = ZSql + "Hasta13 = " + "'" + ZHasta13 + "',"
                ZSql = ZSql + "Desde14 = " + "'" + ZDesde14 + "',"
                ZSql = ZSql + "Hasta14 = " + "'" + ZHasta14 + "',"
                ZSql = ZSql + "Desde15 = " + "'" + ZDesde15 + "',"
                ZSql = ZSql + "Hasta15 = " + "'" + ZHasta15 + "',"
                ZSql = ZSql + "Desde16 = " + "'" + ZDesde16 + "',"
                ZSql = ZSql + "Hasta16 = " + "'" + ZHasta16 + "',"
                ZSql = ZSql + "Desde17 = " + "'" + ZDesde17 + "',"
                ZSql = ZSql + "Hasta17 = " + "'" + ZHasta17 + "',"
                ZSql = ZSql + "Desde18 = " + "'" + ZDesde18 + "',"
                ZSql = ZSql + "Hasta18 = " + "'" + ZHasta18 + "',"
                ZSql = ZSql + "Desde19 = " + "'" + ZDesde19 + "',"
                ZSql = ZSql + "Hasta19 = " + "'" + ZHasta19 + "',"
                ZSql = ZSql + "Desde20 = " + "'" + ZDesde20 + "',"
                ZSql = ZSql + "Hasta20 = " + "'" + ZHasta20 + "'"
                ZSql = ZSql + " Where Clave = " + "'" + ZClave + "'"
                     
                spEspecificacionesUnificaVersion = ZSql
                Set rstEspecificacionesUnificaVersion = db.OpenRecordset(spEspecificacionesUnificaVersion, dbOpenSnapshot, dbSQLPassThrough)
            
            
            
            
            
            
            
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EspecificacionesUnificaVersionII ("
                ZSql = ZSql + "Clave, "
                ZSql = ZSql + "Version, "
                ZSql = ZSql + "Ensayo21, Valor21, "
                ZSql = ZSql + "Ensayo22, Valor22, "
                ZSql = ZSql + "Ensayo23, Valor23, "
                ZSql = ZSql + "Ensayo24, Valor24, "
                ZSql = ZSql + "Ensayo25, Valor25, "
                ZSql = ZSql + "Ensayo26, Valor26, "
                ZSql = ZSql + "Ensayo27, Valor27, "
                ZSql = ZSql + "Ensayo28, Valor28, "
                ZSql = ZSql + "Ensayo29, Valor29, "
                ZSql = ZSql + "Ensayo30, Valor30, "
                ZSql = ZSql + "Desde21, Hasta21, "
                ZSql = ZSql + "Desde22, Hasta22, "
                ZSql = ZSql + "Desde23, Hasta23, "
                ZSql = ZSql + "Desde24, Hasta24, "
                ZSql = ZSql + "Desde25, Hasta25, "
                ZSql = ZSql + "Desde26, Hasta26, "
                ZSql = ZSql + "Desde27, Hasta27, "
                ZSql = ZSql + "Desde28, Hasta28, "
                ZSql = ZSql + "Desde29, Hasta29, "
                ZSql = ZSql + "Desde30, Hasta30, "
                ZSql = ZSql + "Producto) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + ZClave + "',"
                ZSql = ZSql + "'" + ZVersion + "',"
                ZSql = ZSql + "'" + ZEnsayo21 + "'," + "'" + ZValor21 + "',"
                ZSql = ZSql + "'" + ZEnsayo22 + "'," + "'" + ZValor22 + "',"
                ZSql = ZSql + "'" + ZEnsayo23 + "'," + "'" + ZValor23 + "',"
                ZSql = ZSql + "'" + ZEnsayo24 + "'," + "'" + ZValor24 + "',"
                ZSql = ZSql + "'" + ZEnsayo25 + "'," + "'" + ZValor25 + "',"
                ZSql = ZSql + "'" + ZEnsayo26 + "'," + "'" + ZValor26 + "',"
                ZSql = ZSql + "'" + ZEnsayo27 + "'," + "'" + ZValor27 + "',"
                ZSql = ZSql + "'" + ZEnsayo28 + "'," + "'" + ZValor28 + "',"
                ZSql = ZSql + "'" + ZEnsayo29 + "'," + "'" + ZValor29 + "',"
                ZSql = ZSql + "'" + ZEnsayo30 + "'," + "'" + ZValor30 + "',"
                ZSql = ZSql + "'" + ZDesde21 + "'," + "'" + ZHasta21 + "',"
                ZSql = ZSql + "'" + ZDesde22 + "'," + "'" + ZHasta22 + "',"
                ZSql = ZSql + "'" + ZDesde23 + "'," + "'" + ZHasta23 + "',"
                ZSql = ZSql + "'" + ZDesde24 + "'," + "'" + ZHasta24 + "',"
                ZSql = ZSql + "'" + ZDesde25 + "'," + "'" + ZHasta25 + "',"
                ZSql = ZSql + "'" + ZDesde26 + "'," + "'" + ZHasta26 + "',"
                ZSql = ZSql + "'" + ZDesde27 + "'," + "'" + ZHasta27 + "',"
                ZSql = ZSql + "'" + ZDesde28 + "'," + "'" + ZHasta28 + "',"
                ZSql = ZSql + "'" + ZDesde29 + "'," + "'" + ZHasta29 + "',"
                ZSql = ZSql + "'" + ZDesde30 + "'," + "'" + ZHasta30 + "',"
                ZSql = ZSql + "'" + ZProducto + "')"
            
                spEspecificacionesUnificaVersionII = ZSql
                Set rstEspecificacionesUnificaVersionII = db.OpenRecordset(spEspecificacionesUnificaVersionII, dbOpenSnapshot, dbSQLPassThrough)

                Rem
                Rem graba los datos nuevos
                Rem
            
                ZVersion = Str$(Val(Version.Text) + 1)
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
                ZSql = ZSql + "Producto = " + "'" + WProducto + "',"
                ZSql = ZSql + "Ensayo1 = " + "'" + WEnsayo1 + "',"
                ZSql = ZSql + "Valor1 = " + "'" + WValor1 + "',"
                ZSql = ZSql + "Ensayo2 = " + "'" + WEnsayo2 + "',"
                ZSql = ZSql + "Valor2 = " + "'" + WValor2 + "',"
                ZSql = ZSql + "Ensayo3 = " + "'" + WEnsayo3 + "',"
                ZSql = ZSql + "Valor3 = " + "'" + WValor3 + "',"
                ZSql = ZSql + "Ensayo4 = " + "'" + WEnsayo4 + "',"
                ZSql = ZSql + "Valor4 = " + "'" + WValor4 + "',"
                ZSql = ZSql + "Ensayo5 = " + "'" + WEnsayo5 + "',"
                ZSql = ZSql + "Valor5 = " + "'" + WValor5 + "',"
                ZSql = ZSql + "Ensayo6 = " + "'" + WEnsayo6 + "',"
                ZSql = ZSql + "Valor6 = " + "'" + WValor6 + "',"
                ZSql = ZSql + "Ensayo7 = " + "'" + WEnsayo7 + "',"
                ZSql = ZSql + "Valor7 = " + "'" + WValor7 + "',"
                ZSql = ZSql + "Ensayo8 = " + "'" + WEnsayo8 + "',"
                ZSql = ZSql + "Valor8 = " + "'" + WValor8 + "',"
                ZSql = ZSql + "Ensayo9 = " + "'" + WEnsayo9 + "',"
                ZSql = ZSql + "Valor9 = " + "'" + WValor9 + "',"
                ZSql = ZSql + "Ensayo10 = " + "'" + WEnsayo10 + "',"
                ZSql = ZSql + "Valor10 = " + "'" + WValor10 + "',"
                ZSql = ZSql + "Ensayo11 = " + "'" + WEnsayo11 + "',"
                ZSql = ZSql + "Valor11 = " + "'" + WValor11 + "',"
                ZSql = ZSql + "Ensayo12 = " + "'" + WEnsayo12 + "',"
                ZSql = ZSql + "Valor12 = " + "'" + WValor12 + "',"
                ZSql = ZSql + "Ensayo13 = " + "'" + WEnsayo13 + "',"
                ZSql = ZSql + "Valor13 = " + "'" + WValor13 + "',"
                ZSql = ZSql + "Ensayo14 = " + "'" + WEnsayo14 + "',"
                ZSql = ZSql + "Valor14 = " + "'" + WValor14 + "',"
                ZSql = ZSql + "Ensayo15 = " + "'" + WEnsayo15 + "',"
                ZSql = ZSql + "Valor15 = " + "'" + WValor15 + "',"
                ZSql = ZSql + "Ensayo16 = " + "'" + WEnsayo16 + "',"
                ZSql = ZSql + "Valor16 = " + "'" + WValor16 + "',"
                ZSql = ZSql + "Ensayo17 = " + "'" + WEnsayo17 + "',"
                ZSql = ZSql + "Valor17 = " + "'" + WValor17 + "',"
                ZSql = ZSql + "Ensayo18 = " + "'" + WEnsayo18 + "',"
                ZSql = ZSql + "Valor18 = " + "'" + WValor18 + "',"
                ZSql = ZSql + "Ensayo19 = " + "'" + WEnsayo19 + "',"
                ZSql = ZSql + "Valor19 = " + "'" + WValor19 + "',"
                ZSql = ZSql + "Ensayo20 = " + "'" + WEnsayo20 + "',"
                ZSql = ZSql + "Valor20 = " + "'" + WValor20 + "',"
                ZSql = ZSql + "WDate = " + "'" + WDate + "',"
                ZSql = ZSql + "Version = " + "'" + ZVersion + "',"
                ZSql = ZSql + "Fecha = " + "'" + ZFecha + "'"
                ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                     
                spEspecificacionesUnifica = ZSql
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
                        Else
                    
                ZVersion = "1"
                ZFecha = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 4)
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EspecificacionesUnifica ("
                ZSql = ZSql + "Producto, "
                ZSql = ZSql + "Ensayo1, Valor1, "
                ZSql = ZSql + "Ensayo2, Valor2, "
                ZSql = ZSql + "Ensayo3, Valor3, "
                ZSql = ZSql + "Ensayo4, Valor4, "
                ZSql = ZSql + "Ensayo5, Valor5, "
                ZSql = ZSql + "Ensayo6, Valor6, "
                ZSql = ZSql + "Ensayo7, Valor7, "
                ZSql = ZSql + "Ensayo8, Valor8, "
                ZSql = ZSql + "Ensayo9, Valor9, "
                ZSql = ZSql + "Ensayo10, Valor10, "
                ZSql = ZSql + "Ensayo11, Valor11, "
                ZSql = ZSql + "Ensayo12, Valor12, "
                ZSql = ZSql + "Ensayo13, Valor13, "
                ZSql = ZSql + "Ensayo14, Valor14, "
                ZSql = ZSql + "Ensayo15, Valor15, "
                ZSql = ZSql + "Ensayo16, Valor16, "
                ZSql = ZSql + "Ensayo17, Valor17, "
                ZSql = ZSql + "Ensayo18, Valor18, "
                ZSql = ZSql + "Ensayo19, Valor19, "
                ZSql = ZSql + "Ensayo20, Valor20, "
                ZSql = ZSql + "WDate, "
                ZSql = ZSql + "Version, "
                ZSql = ZSql + "Fecha) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WProducto + "',"
                ZSql = ZSql + "'" + WEnsayo1 + "'," + "'" + WValor1 + "',"
                ZSql = ZSql + "'" + WEnsayo2 + "'," + "'" + WValor2 + "',"
                ZSql = ZSql + "'" + WEnsayo3 + "'," + "'" + WValor3 + "',"
                ZSql = ZSql + "'" + WEnsayo4 + "'," + "'" + WValor4 + "',"
                ZSql = ZSql + "'" + WEnsayo5 + "'," + "'" + WValor5 + "',"
                ZSql = ZSql + "'" + WEnsayo6 + "'," + "'" + WValor6 + "',"
                ZSql = ZSql + "'" + WEnsayo7 + "'," + "'" + WValor7 + "',"
                ZSql = ZSql + "'" + WEnsayo8 + "'," + "'" + WValor8 + "',"
                ZSql = ZSql + "'" + WEnsayo9 + "'," + "'" + WValor9 + "',"
                ZSql = ZSql + "'" + WEnsayo10 + "'," + "'" + WValor10 + "',"
                ZSql = ZSql + "'" + WEnsayo11 + "'," + "'" + WValor11 + "',"
                ZSql = ZSql + "'" + WEnsayo12 + "'," + "'" + WValor12 + "',"
                ZSql = ZSql + "'" + WEnsayo13 + "'," + "'" + WValor13 + "',"
                ZSql = ZSql + "'" + WEnsayo14 + "'," + "'" + WValor14 + "',"
                ZSql = ZSql + "'" + WEnsayo15 + "'," + "'" + WValor15 + "',"
                ZSql = ZSql + "'" + WEnsayo16 + "'," + "'" + WValor16 + "',"
                ZSql = ZSql + "'" + WEnsayo17 + "'," + "'" + WValor17 + "',"
                ZSql = ZSql + "'" + WEnsayo18 + "'," + "'" + WValor18 + "',"
                ZSql = ZSql + "'" + WEnsayo19 + "'," + "'" + WValor19 + "',"
                ZSql = ZSql + "'" + WEnsayo20 + "'," + "'" + WValor20 + "',"
                ZSql = ZSql + "'" + WDate + "',"
                ZSql = ZSql + "'" + ZVersion + "',"
                ZSql = ZSql + "'" + ZFecha + "')"
            
                spEspecificacionesUnifica = ZSql
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                
            End If
            
            
            
            
            ZSql = ""
            ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
            ZSql = ZSql + " Operador = " + "'" + ZOperador + "'"
            ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                            
            spEspecificacionesUnifica = ZSql
            Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
        
        
        
        
        
            ZSql = ""
            ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
            ZSql = ZSql + "ControlCambio = " + "'" + ControlCambio.Text + "',"
            ZSql = ZSql + "Desde1 = " + "'" + WDesde1 + "',"
            ZSql = ZSql + "Hasta1 = " + "'" + WHasta1 + "',"
            ZSql = ZSql + "Desde2 = " + "'" + WDesde2 + "',"
            ZSql = ZSql + "Hasta2 = " + "'" + WHasta2 + "',"
            ZSql = ZSql + "Desde3 = " + "'" + WDesde3 + "',"
            ZSql = ZSql + "Hasta3 = " + "'" + WHasta3 + "',"
            ZSql = ZSql + "Desde4 = " + "'" + WDesde4 + "',"
            ZSql = ZSql + "Hasta4 = " + "'" + WHasta4 + "',"
            ZSql = ZSql + "Desde5 = " + "'" + WDesde5 + "',"
            ZSql = ZSql + "Hasta5 = " + "'" + WHasta5 + "',"
            ZSql = ZSql + "Desde6 = " + "'" + WDesde6 + "',"
            ZSql = ZSql + "Hasta6 = " + "'" + WHasta6 + "',"
            ZSql = ZSql + "Desde7 = " + "'" + WDesde7 + "',"
            ZSql = ZSql + "Hasta7 = " + "'" + WHasta7 + "',"
            ZSql = ZSql + "Desde8 = " + "'" + WDesde8 + "',"
            ZSql = ZSql + "Hasta8 = " + "'" + WHasta8 + "',"
            ZSql = ZSql + "Desde9 = " + "'" + WDesde9 + "',"
            ZSql = ZSql + "Hasta9 = " + "'" + WHasta9 + "',"
            ZSql = ZSql + "Desde10 = " + "'" + WDesde10 + "',"
            ZSql = ZSql + "Hasta10 = " + "'" + WHasta10 + "',"
            ZSql = ZSql + "Desde11 = " + "'" + WDesde11 + "',"
            ZSql = ZSql + "Hasta11 = " + "'" + WHasta11 + "',"
            ZSql = ZSql + "Desde12 = " + "'" + WDesde12 + "',"
            ZSql = ZSql + "Hasta12 = " + "'" + WHasta12 + "',"
            ZSql = ZSql + "Desde13 = " + "'" + WDesde13 + "',"
            ZSql = ZSql + "Hasta13 = " + "'" + WHasta13 + "',"
            ZSql = ZSql + "Desde14 = " + "'" + WDesde14 + "',"
            ZSql = ZSql + "Hasta14 = " + "'" + WHasta14 + "',"
            ZSql = ZSql + "Desde15 = " + "'" + WDesde15 + "',"
            ZSql = ZSql + "Hasta15 = " + "'" + WHasta15 + "',"
            ZSql = ZSql + "Desde16 = " + "'" + WDesde16 + "',"
            ZSql = ZSql + "Hasta16 = " + "'" + WHasta16 + "',"
            ZSql = ZSql + "Desde17 = " + "'" + WDesde17 + "',"
            ZSql = ZSql + "Hasta17 = " + "'" + WHasta17 + "',"
            ZSql = ZSql + "Desde18 = " + "'" + WDesde18 + "',"
            ZSql = ZSql + "Hasta18 = " + "'" + WHasta18 + "',"
            ZSql = ZSql + "Desde19 = " + "'" + WDesde19 + "',"
            ZSql = ZSql + "Hasta19 = " + "'" + WHasta19 + "',"
            ZSql = ZSql + "Desde20 = " + "'" + WDesde20 + "',"
            ZSql = ZSql + "Hasta20 = " + "'" + WHasta20 + "'"
            ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                 
            spEspecificacionesUnifica = ZSql
            Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
            
            
            Sql1 = "Select *"
            Sql2 = " FROM EspecificacionesUnificaIII"
            Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Codigo.Text + "'"
            spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnificaIII.RecordCount > 0 Then
            
                rstEspecificacionesUnificaIII.Close
            
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecificacionesUnificaIII SET "
                ZSql = ZSql + "Ensayo21 = " + "'" + WEnsayo21 + "',"
                ZSql = ZSql + "Valor21 = " + "'" + WValor21 + "',"
                ZSql = ZSql + "Desde21 = " + "'" + WDesde21 + "',"
                ZSql = ZSql + "Hasta21 = " + "'" + WHasta21 + "',"
                ZSql = ZSql + "Ensayo22 = " + "'" + WEnsayo22 + "',"
                ZSql = ZSql + "Valor22 = " + "'" + WValor22 + "',"
                ZSql = ZSql + "Desde22 = " + "'" + WDesde22 + "',"
                ZSql = ZSql + "Hasta22 = " + "'" + WHasta22 + "',"
                ZSql = ZSql + "Ensayo23 = " + "'" + WEnsayo23 + "',"
                ZSql = ZSql + "Valor23 = " + "'" + WValor23 + "',"
                ZSql = ZSql + "Desde23 = " + "'" + WDesde23 + "',"
                ZSql = ZSql + "Hasta23 = " + "'" + WHasta23 + "',"
                ZSql = ZSql + "Ensayo24 = " + "'" + WEnsayo24 + "',"
                ZSql = ZSql + "Valor24 = " + "'" + WValor24 + "',"
                ZSql = ZSql + "Desde24 = " + "'" + WDesde24 + "',"
                ZSql = ZSql + "Hasta24 = " + "'" + WHasta24 + "',"
                ZSql = ZSql + "Ensayo25 = " + "'" + WEnsayo25 + "',"
                ZSql = ZSql + "Valor25 = " + "'" + WValor25 + "',"
                ZSql = ZSql + "Desde25 = " + "'" + WDesde25 + "',"
                ZSql = ZSql + "Hasta25 = " + "'" + WHasta25 + "',"
                ZSql = ZSql + "Ensayo26 = " + "'" + WEnsayo26 + "',"
                ZSql = ZSql + "Valor26 = " + "'" + WValor26 + "',"
                ZSql = ZSql + "Desde26 = " + "'" + WDesde26 + "',"
                ZSql = ZSql + "Hasta26 = " + "'" + WHasta26 + "',"
                ZSql = ZSql + "Ensayo27 = " + "'" + WEnsayo27 + "',"
                ZSql = ZSql + "Valor27 = " + "'" + WValor27 + "',"
                ZSql = ZSql + "Desde27 = " + "'" + WDesde27 + "',"
                ZSql = ZSql + "Hasta27 = " + "'" + WHasta27 + "',"
                ZSql = ZSql + "Ensayo28 = " + "'" + WEnsayo28 + "',"
                ZSql = ZSql + "Valor28 = " + "'" + WValor28 + "',"
                ZSql = ZSql + "Desde28 = " + "'" + WDesde28 + "',"
                ZSql = ZSql + "Hasta28 = " + "'" + WHasta28 + "',"
                ZSql = ZSql + "Ensayo29 = " + "'" + WEnsayo29 + "',"
                ZSql = ZSql + "Valor29 = " + "'" + WValor29 + "',"
                ZSql = ZSql + "Desde29 = " + "'" + WDesde29 + "',"
                ZSql = ZSql + "Hasta29 = " + "'" + WHasta29 + "',"
                ZSql = ZSql + "Ensayo30 = " + "'" + WEnsayo30 + "',"
                ZSql = ZSql + "Valor30 = " + "'" + WValor30 + "',"
                ZSql = ZSql + "Desde30 = " + "'" + WDesde30 + "',"
                ZSql = ZSql + "Hasta30 = " + "'" + WHasta30 + "'"
                ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                     
                spEspecificacionesUnificaIII = ZSql
                Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
                
                    Else
                    
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EspecificacionesUnificaIII ("
                ZSql = ZSql + "Producto, "
                ZSql = ZSql + "Ensayo21, "
                ZSql = ZSql + "Valor21, "
                ZSql = ZSql + "Desde21, "
                ZSql = ZSql + "Hasta21, "
                ZSql = ZSql + "Ensayo22, "
                ZSql = ZSql + "Valor22, "
                ZSql = ZSql + "Desde22, "
                ZSql = ZSql + "Hasta22, "
                ZSql = ZSql + "Ensayo23, "
                ZSql = ZSql + "Valor23, "
                ZSql = ZSql + "Desde23, "
                ZSql = ZSql + "Hasta23, "
                ZSql = ZSql + "Ensayo24, "
                ZSql = ZSql + "Valor24, "
                ZSql = ZSql + "Desde24, "
                ZSql = ZSql + "Hasta24, "
                ZSql = ZSql + "Ensayo25, "
                ZSql = ZSql + "Valor25, "
                ZSql = ZSql + "Desde25, "
                ZSql = ZSql + "Hasta25, "
                ZSql = ZSql + "Ensayo26, "
                ZSql = ZSql + "Valor26, "
                ZSql = ZSql + "Desde26, "
                ZSql = ZSql + "Hasta26, "
                ZSql = ZSql + "Ensayo27, "
                ZSql = ZSql + "Valor27, "
                ZSql = ZSql + "Desde27, "
                ZSql = ZSql + "Hasta27, "
                ZSql = ZSql + "Ensayo28, "
                ZSql = ZSql + "Valor28, "
                ZSql = ZSql + "Desde28, "
                ZSql = ZSql + "Hasta28, "
                ZSql = ZSql + "Ensayo29, "
                ZSql = ZSql + "Valor29, "
                ZSql = ZSql + "Desde29, "
                ZSql = ZSql + "Hasta29, "
                ZSql = ZSql + "Ensayo30, "
                ZSql = ZSql + "Valor30, "
                ZSql = ZSql + "Desde30, "
                ZSql = ZSql + "Hasta30) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WProducto + "',"
                ZSql = ZSql + "'" + WEnsayo21 + "',"
                ZSql = ZSql + "'" + WValor21 + "',"
                ZSql = ZSql + "'" + WDesde21 + "',"
                ZSql = ZSql + "'" + WHasta21 + "',"
                ZSql = ZSql + "'" + WEnsayo22 + "',"
                ZSql = ZSql + "'" + WValor22 + "',"
                ZSql = ZSql + "'" + WDesde22 + "',"
                ZSql = ZSql + "'" + WHasta22 + "',"
                ZSql = ZSql + "'" + WEnsayo23 + "',"
                ZSql = ZSql + "'" + WValor23 + "',"
                ZSql = ZSql + "'" + WDesde23 + "',"
                ZSql = ZSql + "'" + WHasta23 + "',"
                ZSql = ZSql + "'" + WEnsayo24 + "',"
                ZSql = ZSql + "'" + WValor24 + "',"
                ZSql = ZSql + "'" + WDesde24 + "',"
                ZSql = ZSql + "'" + WHasta24 + "',"
                ZSql = ZSql + "'" + WEnsayo25 + "',"
                ZSql = ZSql + "'" + WValor25 + "',"
                ZSql = ZSql + "'" + WDesde25 + "',"
                ZSql = ZSql + "'" + WHasta25 + "',"
                ZSql = ZSql + "'" + WEnsayo26 + "',"
                ZSql = ZSql + "'" + WValor26 + "',"
                ZSql = ZSql + "'" + WDesde26 + "',"
                ZSql = ZSql + "'" + WHasta26 + "',"
                ZSql = ZSql + "'" + WEnsayo27 + "',"
                ZSql = ZSql + "'" + WValor27 + "',"
                ZSql = ZSql + "'" + WDesde27 + "',"
                ZSql = ZSql + "'" + WHasta27 + "',"
                ZSql = ZSql + "'" + WEnsayo28 + "',"
                ZSql = ZSql + "'" + WValor28 + "',"
                ZSql = ZSql + "'" + WDesde28 + "',"
                ZSql = ZSql + "'" + WHasta28 + "',"
                ZSql = ZSql + "'" + WEnsayo29 + "',"
                ZSql = ZSql + "'" + WValor29 + "',"
                ZSql = ZSql + "'" + WDesde29 + "',"
                ZSql = ZSql + "'" + WHasta29 + "',"
                ZSql = ZSql + "'" + WEnsayo30 + "',"
                ZSql = ZSql + "'" + WValor30 + "',"
                ZSql = ZSql + "'" + WDesde30 + "',"
                ZSql = ZSql + "'" + WHasta30 + "')"
            
                spEspecificacionesUnificaIII = ZSql
                Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
            
            End If
            
            WProducto = Codigo.Text

            WIValor1 = IValor1.Text
            WIValor2 = IValor2.Text
            WIValor3 = IValor3.Text
            WIValor4 = IValor4.Text
            WIValor5 = IValor5.Text
            WIValor6 = IValor6.Text
            WIValor7 = IValor7.Text
            WIValor8 = IValor8.Text
            WIValor9 = IValor9.Text
            WIValor10 = IValor10.Text
            WIValor11 = IValor11.Text
            WIValor12 = IValor12.Text
            WIValor13 = IValor13.Text
            WIValor14 = IValor14.Text
            WIValor15 = IValor15.Text
            WIValor16 = IValor16.Text
            WIValor17 = IValor17.Text
            WIValor18 = IValor18.Text
            WIValor19 = IValor19.Text
            WIValor20 = IValor20.Text
            WIValor21 = IValor21.Text
            WIValor22 = IValor22.Text
            WIValor23 = IValor23.Text
            WIValor24 = IValor24.Text
            WIValor25 = IValor25.Text
            WIValor26 = IValor26.Text
            WIValor27 = IValor27.Text
            WIValor28 = IValor28.Text
            WIValor29 = IValor29.Text
            WIValor30 = IValor30.Text
            
            Sql1 = "Select EspecificacionesUnificaII.Producto"
            Sql2 = " FROM EspecificacionesUnificaII"
            Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + WProducto + "'"
            spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnificaII.RecordCount > 0 Then
            
                ZSql = ""
                ZSql = ZSql + "UPDATE EspecificacionesUnificaII SET "
                ZSql = ZSql + "DescripcionIngles = " + "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql + "Cas = " + "'" + Cas.Text + "',"
                ZSql = ZSql + "IValor1 = " + "'" + WIValor1 + "',"
                ZSql = ZSql + "IValor2 = " + "'" + WIValor2 + "',"
                ZSql = ZSql + "IValor3 = " + "'" + WIValor3 + "',"
                ZSql = ZSql + "IValor4 = " + "'" + WIValor4 + "',"
                ZSql = ZSql + "IValor5 = " + "'" + WIValor5 + "',"
                ZSql = ZSql + "IValor6 = " + "'" + WIValor6 + "',"
                ZSql = ZSql + "IValor7 = " + "'" + WIValor7 + "',"
                ZSql = ZSql + "IValor8 = " + "'" + WIValor8 + "',"
                ZSql = ZSql + "IValor9 = " + "'" + WIValor9 + "',"
                ZSql = ZSql + "IValor10 = " + "'" + WIValor10 + "',"
                ZSql = ZSql + "IValor11 = " + "'" + WIValor11 + "',"
                ZSql = ZSql + "IValor12 = " + "'" + WIValor12 + "',"
                ZSql = ZSql + "IValor13 = " + "'" + WIValor13 + "',"
                ZSql = ZSql + "IValor14 = " + "'" + WIValor14 + "',"
                ZSql = ZSql + "IValor15 = " + "'" + WIValor15 + "',"
                ZSql = ZSql + "IValor16 = " + "'" + WIValor16 + "',"
                ZSql = ZSql + "IValor17 = " + "'" + WIValor17 + "',"
                ZSql = ZSql + "IValor18 = " + "'" + WIValor18 + "',"
                ZSql = ZSql + "IValor19 = " + "'" + WIValor19 + "',"
                ZSql = ZSql + "IValor20 = " + "'" + WIValor20 + "',"
                ZSql = ZSql + "IValor21 = " + "'" + WIValor21 + "',"
                ZSql = ZSql + "IValor22 = " + "'" + WIValor22 + "',"
                ZSql = ZSql + "IValor23 = " + "'" + WIValor23 + "',"
                ZSql = ZSql + "IValor24 = " + "'" + WIValor24 + "',"
                ZSql = ZSql + "IValor25 = " + "'" + WIValor25 + "',"
                ZSql = ZSql + "IValor26 = " + "'" + WIValor26 + "',"
                ZSql = ZSql + "IValor27 = " + "'" + WIValor27 + "',"
                ZSql = ZSql + "IValor28 = " + "'" + WIValor28 + "',"
                ZSql = ZSql + "IValor29 = " + "'" + WIValor29 + "',"
                ZSql = ZSql + "IValor30 = " + "'" + WIValor30 + "'"
                ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                     
                spEspecificacionesUnificaII = ZSql
                Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
            
            
                    Else
                
                ZSql = ""
                ZSql = ZSql + "INSERT INTO EspecificacionesUnificaII ("
                ZSql = ZSql + "Producto, "
                ZSql = ZSql + "DescripcionIngles, "
                ZSql = ZSql + "Cas, "
                ZSql = ZSql + "IValor1, "
                ZSql = ZSql + "IValor2, "
                ZSql = ZSql + "IValor3, "
                ZSql = ZSql + "IValor4, "
                ZSql = ZSql + "IValor5, "
                ZSql = ZSql + "IValor6, "
                ZSql = ZSql + "IValor7, "
                ZSql = ZSql + "IValor8, "
                ZSql = ZSql + "IValor9, "
                ZSql = ZSql + "IValor10, "
                ZSql = ZSql + "IValor11, "
                ZSql = ZSql + "IValor12, "
                ZSql = ZSql + "IValor13, "
                ZSql = ZSql + "IValor14, "
                ZSql = ZSql + "IValor15, "
                ZSql = ZSql + "IValor16, "
                ZSql = ZSql + "IValor17, "
                ZSql = ZSql + "IValor18, "
                ZSql = ZSql + "IValor19, "
                ZSql = ZSql + "IValor20, "
                ZSql = ZSql + "IValor21, "
                ZSql = ZSql + "IValor22, "
                ZSql = ZSql + "IValor23, "
                ZSql = ZSql + "IValor24, "
                ZSql = ZSql + "IValor25, "
                ZSql = ZSql + "IValor26, "
                ZSql = ZSql + "IValor27, "
                ZSql = ZSql + "IValor28, "
                ZSql = ZSql + "IValor29, "
                ZSql = ZSql + "Ivalor30) "
                ZSql = ZSql + "Values ("
                ZSql = ZSql + "'" + WProducto + "',"
                ZSql = ZSql + "'" + DescripcionIngles.Text + "',"
                ZSql = ZSql + "'" + Cas.Text + "',"
                ZSql = ZSql + "'" + WIValor1 + "',"
                ZSql = ZSql + "'" + WIValor2 + "',"
                ZSql = ZSql + "'" + WIValor3 + "',"
                ZSql = ZSql + "'" + WIValor4 + "',"
                ZSql = ZSql + "'" + WIValor5 + "',"
                ZSql = ZSql + "'" + WIValor6 + "',"
                ZSql = ZSql + "'" + WIValor7 + "',"
                ZSql = ZSql + "'" + WIValor8 + "',"
                ZSql = ZSql + "'" + WIValor9 + "',"
                ZSql = ZSql + "'" + WIValor10 + "',"
                ZSql = ZSql + "'" + WIValor11 + "',"
                ZSql = ZSql + "'" + WIValor12 + "',"
                ZSql = ZSql + "'" + WIValor13 + "',"
                ZSql = ZSql + "'" + WIValor14 + "',"
                ZSql = ZSql + "'" + WIValor15 + "',"
                ZSql = ZSql + "'" + WIValor16 + "',"
                ZSql = ZSql + "'" + WIValor17 + "',"
                ZSql = ZSql + "'" + WIValor18 + "',"
                ZSql = ZSql + "'" + WIValor19 + "',"
                ZSql = ZSql + "'" + WIValor20 + "',"
                ZSql = ZSql + "'" + WIValor21 + "',"
                ZSql = ZSql + "'" + WIValor22 + "',"
                ZSql = ZSql + "'" + WIValor23 + "',"
                ZSql = ZSql + "'" + WIValor24 + "',"
                ZSql = ZSql + "'" + WIValor25 + "',"
                ZSql = ZSql + "'" + WIValor26 + "',"
                ZSql = ZSql + "'" + WIValor27 + "',"
                ZSql = ZSql + "'" + WIValor28 + "',"
                ZSql = ZSql + "'" + WIValor29 + "',"
                ZSql = ZSql + "'" + WIValor30 + "')"
            
                spEspecificacionesUnificaII = ZSql
                Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)

            End If
            
            Call Conecta_Empresa
            
            Call ImprimeAutomatico
        
            Call CmdLimpiar_Click
            Codigo.SetFocus
            
        End If
        
    End If
    
End Sub

Private Sub cmdDelete_Click()
    If Codigo.Text <> "" Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        Sql1 = "Select *"
        Sql2 = " FROM EspecificacionesUnifica"
        Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
        spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnifica.RecordCount > 0 Then
            rstEspecificacionesUnifica.Close
            T$ = "Borrar Registro"
            m$ = "Desea Borrar el Registro "
            Respuesta% = MsgBox(m$, 32 + 4, T$)
            If Respuesta% = 6 Then
                Sql1 = "DELETE EspecificacionesUnifica"
                Sql2 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
                spEspecificacionesUnifica = Sql1 + Sql2
                Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
                Call CmdLimpiar_Click
            End If
        End If
        
        Call Conecta_Empresa
        
    End If
    Codigo.SetFocus
End Sub

Private Sub CmdLimpiar_Click()
    Codigo.Text = "  -   -   "
    Ensayo1.Text = ""
    Valor1.Text = ""
    Ensayo2.Text = ""
    valor2.Text = ""
    Ensayo3.Text = ""
    Valor3.Text = ""
    Ensayo4.Text = ""
    valor4.Text = ""
    Ensayo5.Text = ""
    valor5.Text = ""
    Ensayo6.Text = ""
    valor6.Text = ""
    Ensayo7.Text = ""
    valor7.Text = ""
    Ensayo8.Text = ""
    valor8.Text = ""
    Ensayo9.Text = ""
    valor9.Text = ""
    Ensayo10.Text = ""
    valor10.Text = ""
    Ensayo11.Text = ""
    Valor11.Text = ""
    Ensayo12.Text = ""
    Valor12.Text = ""
    Ensayo13.Text = ""
    Valor13.Text = ""
    Ensayo14.Text = ""
    Valor14.Text = ""
    Ensayo15.Text = ""
    Valor15.Text = ""
    Ensayo16.Text = ""
    Valor16.Text = ""
    Ensayo17.Text = ""
    Valor17.Text = ""
    Ensayo18.Text = ""
    Valor18.Text = ""
    Ensayo19.Text = ""
    Valor19.Text = ""
    Ensayo20.Text = ""
    Valor20.Text = ""
    Ensayo21.Text = ""
    Valor21.Text = ""
    Ensayo22.Text = ""
    Valor22.Text = ""
    Ensayo23.Text = ""
    Valor23.Text = ""
    Ensayo24.Text = ""
    Valor24.Text = ""
    Ensayo25.Text = ""
    Valor25.Text = ""
    Ensayo26.Text = ""
    Valor26.Text = ""
    Ensayo27.Text = ""
    Valor27.Text = ""
    Ensayo28.Text = ""
    Valor28.Text = ""
    Ensayo29.Text = ""
    Valor29.Text = ""
    Ensayo30.Text = ""
    Valor30.Text = ""
    
    IValor1.Text = ""
    IValor2.Text = ""
    IValor3.Text = ""
    IValor4.Text = ""
    IValor5.Text = ""
    IValor6.Text = ""
    IValor7.Text = ""
    IValor8.Text = ""
    IValor9.Text = ""
    IValor10.Text = ""
    IValor11.Text = ""
    IValor12.Text = ""
    IValor13.Text = ""
    IValor14.Text = ""
    IValor15.Text = ""
    IValor16.Text = ""
    IValor17.Text = ""
    IValor18.Text = ""
    IValor19.Text = ""
    IValor20.Text = ""
    IValor21.Text = ""
    IValor22.Text = ""
    IValor23.Text = ""
    IValor24.Text = ""
    IValor25.Text = ""
    IValor26.Text = ""
    IValor27.Text = ""
    IValor28.Text = ""
    IValor29.Text = ""
    IValor30.Text = ""
    
    
    Desde1.Text = ""
    Hasta1.Text = ""
    Desde2.Text = ""
    Hasta2.Text = ""
    Desde3.Text = ""
    Hasta3.Text = ""
    Desde4.Text = ""
    Hasta4.Text = ""
    Desde5.Text = ""
    Hasta5.Text = ""
    Desde6.Text = ""
    Hasta6.Text = ""
    Desde7.Text = ""
    Hasta7.Text = ""
    Desde8.Text = ""
    Hasta8.Text = ""
    Desde9.Text = ""
    Hasta9.Text = ""
    Desde10.Text = ""
    Hasta10.Text = ""
    Desde11.Text = ""
    Hasta11.Text = ""
    Desde12.Text = ""
    Hasta12.Text = ""
    Desde13.Text = ""
    Hasta13.Text = ""
    Desde14.Text = ""
    Hasta14.Text = ""
    Desde15.Text = ""
    Hasta15.Text = ""
    Desde16.Text = ""
    Hasta16.Text = ""
    Desde17.Text = ""
    Hasta17.Text = ""
    Desde18.Text = ""
    Hasta18.Text = ""
    Desde19.Text = ""
    Hasta19.Text = ""
    Desde20.Text = ""
    Hasta20.Text = ""
    Desde21.Text = ""
    Hasta21.Text = ""
    Desde22.Text = ""
    Hasta22.Text = ""
    Desde23.Text = ""
    Hasta23.Text = ""
    Desde24.Text = ""
    Hasta24.Text = ""
    Desde25.Text = ""
    Hasta25.Text = ""
    Desde26.Text = ""
    Hasta26.Text = ""
    Desde27.Text = ""
    Hasta27.Text = ""
    Desde28.Text = ""
    Hasta28.Text = ""
    Desde29.Text = ""
    Hasta29.Text = ""
    Desde30.Text = ""
    Hasta30.Text = ""
    
    Descriprod.Caption = ""
    Descri1.Caption = ""
    descri2.Caption = ""
    Descri3.Caption = ""
    Descri4.Caption = ""
    Descri5.Caption = ""
    Descri6.Caption = ""
    Descri7.Caption = ""
    Descri8.Caption = ""
    Descri9.Caption = ""
    Descri10.Caption = ""
    Descri11.Caption = ""
    Descri12.Caption = ""
    Descri13.Caption = ""
    Descri14.Caption = ""
    Descri15.Caption = ""
    Descri16.Caption = ""
    Descri17.Caption = ""
    Descri18.Caption = ""
    Descri19.Caption = ""
    Descri20.Caption = ""
    Descri21.Caption = ""
    Descri22.Caption = ""
    Descri23.Caption = ""
    Descri24.Caption = ""
    Descri25.Caption = ""
    Descri26.Caption = ""
    Descri27.Caption = ""
    Descri28.Caption = ""
    Descri29.Caption = ""
    Descri30.Caption = ""
    
    Version.Text = ""
    fecha.Text = ""
    ControlCambio.Text = ""
    WGraba = ""
    
    SSTab1.Tab = 0
    Titulo.Caption = "Valor Standard"
    TituloII.Caption = "Valor Standard"
    TituloIII.Caption = "Valor Standard"
    
    Valor1.Visible = True
    valor2.Visible = True
    Valor3.Visible = True
    valor4.Visible = True
    valor5.Visible = True
    valor6.Visible = True
    valor7.Visible = True
    valor8.Visible = True
    valor9.Visible = True
    valor10.Visible = True
    Valor11.Visible = True
    Valor12.Visible = True
    Valor13.Visible = True
    Valor14.Visible = True
    Valor15.Visible = True
    Valor16.Visible = True
    Valor17.Visible = True
    Valor18.Visible = True
    Valor19.Visible = True
    Valor20.Visible = True
    Valor21.Visible = True
    Valor22.Visible = True
    Valor23.Visible = True
    Valor24.Visible = True
    Valor25.Visible = True
    Valor26.Visible = True
    Valor27.Visible = True
    Valor28.Visible = True
    Valor29.Visible = True
    Valor30.Visible = True
    
    IValor1.Visible = False
    IValor2.Visible = False
    IValor3.Visible = False
    IValor4.Visible = False
    IValor5.Visible = False
    IValor6.Visible = False
    IValor7.Visible = False
    IValor8.Visible = False
    IValor9.Visible = False
    IValor10.Visible = False
    IValor11.Visible = False
    IValor12.Visible = False
    IValor13.Visible = False
    IValor14.Visible = False
    IValor15.Visible = False
    IValor16.Visible = False
    IValor17.Visible = False
    IValor18.Visible = False
    IValor19.Visible = False
    IValor20.Visible = False
    IValor21.Visible = False
    IValor22.Visible = False
    IValor23.Visible = False
    IValor24.Visible = False
    IValor25.Visible = False
    IValor26.Visible = False
    IValor27.Visible = False
    IValor28.Visible = False
    IValor29.Visible = False
    IValor30.Visible = False
    
    DescripcionIngles.Text = ""
    Cas.Text = ""
    
    
    Codigo.SetFocus
End Sub

Private Sub cmdClose_Click()
    PrgEspecifiUnifica.Hide
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Command1_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
        
    ZVersion = "1"
    ZFecha = "01/01/2004"
    
    Sql1 = "UPDATE EspecificacionesUnifica SET "
    Sql2 = "Version = " + "'" + ZVersion + "',"
    Sql3 = "Fecha = " + "'" + ZFecha + "'"
                     
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
        
    Call Conecta_Empresa

End Sub

Private Sub Command2_Click()
    Call ImprimeAutomatico
End Sub



Private Sub Ensayo1_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo1.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri1.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor1.SetFocus
                    Else
                IValor1.SetFocus
            End If
                    Else
            Descri1.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo2_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo2.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            descri2.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor2.SetFocus
                    Else
                IValor2.SetFocus
            End If
                    Else
            descri2.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo3_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo3.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri3.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor3.SetFocus
                    Else
                IValor3.SetFocus
            End If
                    Else
            Descri3.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo4_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo4.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri4.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor4.SetFocus
                    Else
                IValor4.SetFocus
            End If
                    Else
            Descri4.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo5_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo5.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri5.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor5.SetFocus
                    Else
                IValor5.SetFocus
            End If
                    Else
            Descri5.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo6_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo6.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri6.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor6.SetFocus
                    Else
                IValor6.SetFocus
            End If
                    Else
            Descri6.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo7_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo7.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri7.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor7.SetFocus
                    Else
                IValor7.SetFocus
            End If
                    Else
            Descri7.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo8_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo8.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri8.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor8.SetFocus
                    Else
                IValor8.SetFocus
            End If
                    Else
            Descri8.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo9_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo9.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri9.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor9.SetFocus
                    Else
                IValor9.SetFocus
            End If
                    Else
            Descri9.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo10_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo10.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri10.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                valor10.SetFocus
                    Else
                IValor10.SetFocus
            End If
                    Else
            Descri10.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo11_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo11.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri11.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor11.SetFocus
                    Else
                IValor11.SetFocus
            End If
                    Else
            Descri11.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo12_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo12.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri12.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor12.SetFocus
                    Else
                IValor12.SetFocus
            End If
                    Else
            Descri12.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo13_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo13.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri13.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor13.SetFocus
                    Else
                IValor13.SetFocus
            End If
                    Else
            Descri13.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo14_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo14.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri14.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor14.SetFocus
                    Else
                IValor14.SetFocus
            End If
                    Else
            Descri14.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo15_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo15.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri15.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor15.SetFocus
                    Else
                IValor15.SetFocus
            End If
                    Else
            Descri15.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo16_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo16.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri16.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor16.SetFocus
                    Else
                IValor16.SetFocus
            End If
                    Else
            Descri16.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo17_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo17.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri17.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor17.SetFocus
                    Else
                IValor17.SetFocus
            End If
                    Else
            Descri17.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo18_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo18.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri18.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor18.SetFocus
                    Else
                IValor18.SetFocus
            End If
                    Else
            Descri18.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo19_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo19.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri19.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor19.SetFocus
                    Else
                IValor19.SetFocus
            End If
                    Else
            Descri19.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo20_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo20.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri20.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor20.SetFocus
                    Else
                IValor20.SetFocus
            End If
                    Else
            Descri20.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo21_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo21.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri21.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor21.SetFocus
                    Else
                IValor21.SetFocus
            End If
                    Else
            Descri21.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo22_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo22.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri22.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor22.SetFocus
                    Else
                IValor22.SetFocus
            End If
                    Else
            Descri22.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo23_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo23.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri23.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor23.SetFocus
                    Else
                IValor23.SetFocus
            End If
                    Else
            Descri23.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo24_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo24.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri24.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor24.SetFocus
                    Else
                IValor24.SetFocus
            End If
                    Else
            Descri24.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo25_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo25.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri25.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor25.SetFocus
                    Else
                IValor25.SetFocus
            End If
                    Else
            Descri25.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo26_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo26.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri26.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor26.SetFocus
                    Else
                IValor26.SetFocus
            End If
                    Else
            Descri26.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo27_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo27.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri27.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor27.SetFocus
                    Else
                IValor27.SetFocus
            End If
                    Else
            Descri27.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo28_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo28.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri28.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor28.SetFocus
                    Else
                IValor28.SetFocus
            End If
                    Else
            Descri28.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo29_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo29.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri29.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor29.SetFocus
                    Else
                IValor29.SetFocus
            End If
                    Else
            Descri29.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Ensayo30_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
        
        spEnsayo = "ConsultaEnsayos " + "'" + Ensayo30.Text + "'"
        Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
        If rstEnsayo.RecordCount > 0 Then
            Descri30.Caption = rstEnsayo!Descripcion
            rstEnsayo.Close
            If Titulo.Caption = "Valor Standard" Then
                Valor30.SetFocus
                    Else
                IValor30.SetFocus
            End If
                    Else
            Descri30.Caption = ""
        End If
        
        Call Conecta_Empresa
        
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Form_Activate()
    Select Case Val(EmpresaActual)
        Case 1
            WEmpresa = "0001"
            txtOdbc = "Empresa01"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2
            WEmpresa = "0002"
            txtOdbc = "Empresa02"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 3
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 4
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 5
            WEmpresa = "0005"
            txtOdbc = "Empresa05"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 6
            WEmpresa = "0006"
            txtOdbc = "Empresa06"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 7
            WEmpresa = "0007"
            txtOdbc = "Empresa07"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 8
            WEmpresa = "0008"
            txtOdbc = "Empresa08"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 9
            WEmpresa = "0009"
            txtOdbc = "Empresa09"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 10
            WEmpresa = "0010"
            txtOdbc = "Empresa10"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 11
            WEmpresa = "0011"
            txtOdbc = "Empresa11"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
    OPEN_FILE_Empresa
End Sub

Private Sub Form_Load()
    With rstEmpresa
        .Index = "Empresa"
        .Seek "=", Val(WEmpresa)
        If .NoMatch = False Then
            PrgEspecifiUnifica.Caption = "Ingreso de Especificaciones de Materia Prima (Unificado) :  " + !Nombre
        End If
    End With
    
    Titulo.Caption = "Valor Standard"
    TituloII.Caption = "Valor Standard"
    TituloIII.Caption = "Valor Standard"
    
    Valor1.Visible = True
    valor2.Visible = True
    Valor3.Visible = True
    valor4.Visible = True
    valor5.Visible = True
    valor6.Visible = True
    valor7.Visible = True
    valor8.Visible = True
    valor9.Visible = True
    valor10.Visible = True
    Valor11.Visible = True
    Valor12.Visible = True
    Valor13.Visible = True
    Valor14.Visible = True
    Valor15.Visible = True
    Valor16.Visible = True
    Valor17.Visible = True
    Valor18.Visible = True
    Valor19.Visible = True
    Valor20.Visible = True
    Valor21.Visible = True
    Valor22.Visible = True
    Valor23.Visible = True
    Valor24.Visible = True
    Valor25.Visible = True
    Valor26.Visible = True
    Valor27.Visible = True
    Valor28.Visible = True
    Valor29.Visible = True
    Valor30.Visible = True
    
    IValor1.Visible = False
    IValor2.Visible = False
    IValor3.Visible = False
    IValor4.Visible = False
    IValor5.Visible = False
    IValor6.Visible = False
    IValor7.Visible = False
    IValor8.Visible = False
    IValor9.Visible = False
    IValor10.Visible = False
    IValor11.Visible = False
    IValor12.Visible = False
    IValor13.Visible = False
    IValor14.Visible = False
    IValor15.Visible = False
    IValor16.Visible = False
    IValor17.Visible = False
    IValor18.Visible = False
    IValor19.Visible = False
    IValor20.Visible = False
    IValor21.Visible = False
    IValor22.Visible = False
    IValor23.Visible = False
    IValor24.Visible = False
    IValor25.Visible = False
    IValor26.Visible = False
    IValor27.Visible = False
    IValor28.Visible = False
    IValor29.Visible = False
    IValor30.Visible = False
    
    DescripcionIngles.Text = ""
    Cas.Text = ""
    
    DesOperador.Caption = ""
    WGraba = ""
    
    EmpresaActual = WEmpresa
    
End Sub

Private Sub GrabaII_Click()

    If WGraba <> "S" Then
    
        ZZProceso = 1
        Call Ingresa_clave

               Else

    
        XEmpresa = WEmpresa
        Select Case Val(WEmpresa)
            Case 1, 3, 5, 6, 7, 10, 11
                WEmpresa = "0003"
                txtOdbc = "Empresa03"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            Case Else
                WEmpresa = "0004"
                txtOdbc = "Empresa04"
                strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        End Select
    
        
        WProducto = Codigo.Text
    
        WIValor1 = IValor1.Text
        WIValor2 = IValor2.Text
        WIValor3 = IValor3.Text
        WIValor4 = IValor4.Text
        WIValor5 = IValor5.Text
        WIValor6 = IValor6.Text
        WIValor7 = IValor7.Text
        WIValor8 = IValor8.Text
        WIValor9 = IValor9.Text
        WIValor10 = IValor10.Text
        WIValor11 = IValor11.Text
        WIValor12 = IValor12.Text
        WIValor13 = IValor13.Text
        WIValor14 = IValor14.Text
        WIValor15 = IValor15.Text
        WIValor16 = IValor16.Text
        WIValor17 = IValor17.Text
        WIValor18 = IValor18.Text
        WIValor19 = IValor19.Text
        WIValor20 = IValor20.Text
        WIValor21 = IValor21.Text
        WIValor22 = IValor22.Text
        WIValor23 = IValor23.Text
        WIValor24 = IValor24.Text
        WIValor25 = IValor25.Text
        WIValor26 = IValor26.Text
        WIValor27 = IValor27.Text
        WIValor28 = IValor28.Text
        WIValor29 = IValor29.Text
        WIValor30 = IValor30.Text
        
        Sql1 = "Select EspecificacionesUnificaII.Producto"
        Sql2 = " FROM EspecificacionesUnificaII"
        Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + WProducto + "'"
        spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
        Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
        If rstEspecificacionesUnificaII.RecordCount > 0 Then
        
            rstEspecificacionesUnificaII.Close
        
            ZSql = ""
            ZSql = ZSql + "UPDATE EspecificacionesUnificaII SET "
            ZSql = ZSql + "DescripcionIngles = " + "'" + DescripcionIngles.Text + "',"
            ZSql = ZSql + "Cas = " + "'" + Cas.Text + "',"
            ZSql = ZSql + "IValor1 = " + "'" + WIValor1 + "',"
            ZSql = ZSql + "IValor2 = " + "'" + WIValor2 + "',"
            ZSql = ZSql + "IValor3 = " + "'" + WIValor3 + "',"
            ZSql = ZSql + "IValor4 = " + "'" + WIValor4 + "',"
            ZSql = ZSql + "IValor5 = " + "'" + WIValor5 + "',"
            ZSql = ZSql + "IValor6 = " + "'" + WIValor6 + "',"
            ZSql = ZSql + "IValor7 = " + "'" + WIValor7 + "',"
            ZSql = ZSql + "IValor8 = " + "'" + WIValor8 + "',"
            ZSql = ZSql + "IValor9 = " + "'" + WIValor9 + "',"
            ZSql = ZSql + "IValor10 = " + "'" + WIValor10 + "',"
            ZSql = ZSql + "IValor11 = " + "'" + WIValor11 + "',"
            ZSql = ZSql + "IValor12 = " + "'" + WIValor12 + "',"
            ZSql = ZSql + "IValor13 = " + "'" + WIValor13 + "',"
            ZSql = ZSql + "IValor14 = " + "'" + WIValor14 + "',"
            ZSql = ZSql + "IValor15 = " + "'" + WIValor15 + "',"
            ZSql = ZSql + "IValor16 = " + "'" + WIValor16 + "',"
            ZSql = ZSql + "IValor17 = " + "'" + WIValor17 + "',"
            ZSql = ZSql + "IValor18 = " + "'" + WIValor18 + "',"
            ZSql = ZSql + "IValor19 = " + "'" + WIValor19 + "',"
            ZSql = ZSql + "IValor20 = " + "'" + WIValor20 + "',"
            ZSql = ZSql + "IValor21 = " + "'" + WIValor21 + "',"
            ZSql = ZSql + "IValor22 = " + "'" + WIValor22 + "',"
            ZSql = ZSql + "IValor23 = " + "'" + WIValor23 + "',"
            ZSql = ZSql + "IValor24 = " + "'" + WIValor24 + "',"
            ZSql = ZSql + "IValor25 = " + "'" + WIValor25 + "',"
            ZSql = ZSql + "IValor26 = " + "'" + WIValor26 + "',"
            ZSql = ZSql + "IValor27 = " + "'" + WIValor27 + "',"
            ZSql = ZSql + "IValor28 = " + "'" + WIValor28 + "',"
            ZSql = ZSql + "IValor29 = " + "'" + WIValor29 + "',"
            ZSql = ZSql + "IValor30 = " + "'" + WIValor30 + "'"
            ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
                 
            spEspecificacionesUnificaII = ZSql
            Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
        
                Else
            
            ZSql = ""
            ZSql = ZSql + "INSERT INTO EspecificacionesUnificaII ("
            ZSql = ZSql + "Producto, "
            ZSql = ZSql + "DescripcionIngles, "
            ZSql = ZSql + "Cas, "
            ZSql = ZSql + "IValor1, "
            ZSql = ZSql + "IValor2, "
            ZSql = ZSql + "IValor3, "
            ZSql = ZSql + "IValor4, "
            ZSql = ZSql + "IValor5, "
            ZSql = ZSql + "IValor6, "
            ZSql = ZSql + "IValor7, "
            ZSql = ZSql + "IValor8, "
            ZSql = ZSql + "IValor9, "
            ZSql = ZSql + "IValor10, "
            ZSql = ZSql + "IValor11, "
            ZSql = ZSql + "IValor12, "
            ZSql = ZSql + "IValor13, "
            ZSql = ZSql + "IValor14, "
            ZSql = ZSql + "IValor15, "
            ZSql = ZSql + "IValor16, "
            ZSql = ZSql + "IValor17, "
            ZSql = ZSql + "IValor18, "
            ZSql = ZSql + "IValor19, "
            ZSql = ZSql + "IValor20, "
            ZSql = ZSql + "IValor21, "
            ZSql = ZSql + "IValor22, "
            ZSql = ZSql + "IValor23, "
            ZSql = ZSql + "IValor24, "
            ZSql = ZSql + "IValor25, "
            ZSql = ZSql + "IValor26, "
            ZSql = ZSql + "IValor27, "
            ZSql = ZSql + "IValor28, "
            ZSql = ZSql + "IValor29, "
            ZSql = ZSql + "Ivalor30) "
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + WProducto + "',"
            ZSql = ZSql + "'" + DescripcionIngles.Text + "',"
            ZSql = ZSql + "'" + Cas.Text + "',"
            ZSql = ZSql + "'" + WIValor1 + "',"
            ZSql = ZSql + "'" + WIValor2 + "',"
            ZSql = ZSql + "'" + WIValor3 + "',"
            ZSql = ZSql + "'" + WIValor4 + "',"
            ZSql = ZSql + "'" + WIValor5 + "',"
            ZSql = ZSql + "'" + WIValor6 + "',"
            ZSql = ZSql + "'" + WIValor7 + "',"
            ZSql = ZSql + "'" + WIValor8 + "',"
            ZSql = ZSql + "'" + WIValor9 + "',"
            ZSql = ZSql + "'" + WIValor10 + "',"
            ZSql = ZSql + "'" + WIValor11 + "',"
            ZSql = ZSql + "'" + WIValor12 + "',"
            ZSql = ZSql + "'" + WIValor13 + "',"
            ZSql = ZSql + "'" + WIValor14 + "',"
            ZSql = ZSql + "'" + WIValor15 + "',"
            ZSql = ZSql + "'" + WIValor16 + "',"
            ZSql = ZSql + "'" + WIValor17 + "',"
            ZSql = ZSql + "'" + WIValor18 + "',"
            ZSql = ZSql + "'" + WIValor19 + "',"
            ZSql = ZSql + "'" + WIValor20 + "',"
            ZSql = ZSql + "'" + WIValor21 + "',"
            ZSql = ZSql + "'" + WIValor22 + "',"
            ZSql = ZSql + "'" + WIValor23 + "',"
            ZSql = ZSql + "'" + WIValor24 + "',"
            ZSql = ZSql + "'" + WIValor25 + "',"
            ZSql = ZSql + "'" + WIValor26 + "',"
            ZSql = ZSql + "'" + WIValor27 + "',"
            ZSql = ZSql + "'" + WIValor28 + "',"
            ZSql = ZSql + "'" + WIValor29 + "',"
            ZSql = ZSql + "'" + WIValor30 + "')"
        
            spEspecificacionesUnificaII = ZSql
            Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
    
        End If
    
        WEnsayo1 = Ensayo1.Text
        WEnsayo2 = Ensayo2.Text
        WEnsayo3 = Ensayo3.Text
        WEnsayo4 = Ensayo4.Text
        WEnsayo5 = Ensayo5.Text
        WEnsayo6 = Ensayo6.Text
        WEnsayo7 = Ensayo7.Text
        WEnsayo8 = Ensayo8.Text
        WEnsayo9 = Ensayo9.Text
        WEnsayo10 = Ensayo10.Text
        WEnsayo11 = Ensayo11.Text
        WEnsayo12 = Ensayo12.Text
        WEnsayo13 = Ensayo13.Text
        WEnsayo14 = Ensayo14.Text
        WEnsayo15 = Ensayo15.Text
        WEnsayo16 = Ensayo16.Text
        WEnsayo17 = Ensayo17.Text
        WEnsayo18 = Ensayo18.Text
        WEnsayo19 = Ensayo19.Text
        WEnsayo20 = Ensayo20.Text
        WEnsayo21 = Ensayo21.Text
        WEnsayo22 = Ensayo22.Text
        WEnsayo23 = Ensayo23.Text
        WEnsayo24 = Ensayo24.Text
        WEnsayo25 = Ensayo25.Text
        WEnsayo26 = Ensayo26.Text
        WEnsayo27 = Ensayo27.Text
        WEnsayo28 = Ensayo28.Text
        WEnsayo29 = Ensayo29.Text
        WEnsayo30 = Ensayo30.Text
    
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecificacionesUnifica SET "
        ZSql = ZSql + "Ensayo1 = " + "'" + WEnsayo1 + "',"
        ZSql = ZSql + "Ensayo2 = " + "'" + WEnsayo2 + "',"
        ZSql = ZSql + "Ensayo3 = " + "'" + WEnsayo3 + "',"
        ZSql = ZSql + "Ensayo4 = " + "'" + WEnsayo4 + "',"
        ZSql = ZSql + "Ensayo5 = " + "'" + WEnsayo5 + "',"
        ZSql = ZSql + "Ensayo6 = " + "'" + WEnsayo6 + "',"
        ZSql = ZSql + "Ensayo7 = " + "'" + WEnsayo7 + "',"
        ZSql = ZSql + "Ensayo8 = " + "'" + WEnsayo8 + "',"
        ZSql = ZSql + "Ensayo9 = " + "'" + WEnsayo9 + "',"
        ZSql = ZSql + "Ensayo10 = " + "'" + WEnsayo10 + "',"
        ZSql = ZSql + "Ensayo11 = " + "'" + WEnsayo11 + "',"
        ZSql = ZSql + "Ensayo12 = " + "'" + WEnsayo12 + "',"
        ZSql = ZSql + "Ensayo13 = " + "'" + WEnsayo13 + "',"
        ZSql = ZSql + "Ensayo14 = " + "'" + WEnsayo14 + "',"
        ZSql = ZSql + "Ensayo15 = " + "'" + WEnsayo15 + "',"
        ZSql = ZSql + "Ensayo16 = " + "'" + WEnsayo16 + "',"
        ZSql = ZSql + "Ensayo17 = " + "'" + WEnsayo17 + "',"
        ZSql = ZSql + "Ensayo18 = " + "'" + WEnsayo18 + "',"
        ZSql = ZSql + "Ensayo19 = " + "'" + WEnsayo19 + "',"
        ZSql = ZSql + "Ensayo20 = " + "'" + WEnsayo20 + "'"
        ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
             
        spEspecificacionesUnifica = ZSql
        Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    
    
    
        ZSql = ""
        ZSql = ZSql + "UPDATE EspecificacionesUnificaIII SET "
        ZSql = ZSql + "Ensayo21 = " + "'" + WEnsayo21 + "',"
        ZSql = ZSql + "Ensayo22 = " + "'" + WEnsayo22 + "',"
        ZSql = ZSql + "Ensayo23 = " + "'" + WEnsayo23 + "',"
        ZSql = ZSql + "Ensayo24 = " + "'" + WEnsayo24 + "',"
        ZSql = ZSql + "Ensayo25 = " + "'" + WEnsayo25 + "',"
        ZSql = ZSql + "Ensayo26 = " + "'" + WEnsayo26 + "',"
        ZSql = ZSql + "Ensayo27 = " + "'" + WEnsayo27 + "',"
        ZSql = ZSql + "Ensayo28 = " + "'" + WEnsayo28 + "',"
        ZSql = ZSql + "Ensayo29 = " + "'" + WEnsayo29 + "',"
        ZSql = ZSql + "Ensayo30 = " + "'" + WEnsayo30 + "'"
        ZSql = ZSql + " Where Producto = " + "'" + WProducto + "'"
             
        spEspecificacionesUnificaIII = ZSql
        Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    
    
        Call Conecta_Empresa
        Call CmdLimpiar_Click
        
    End If
        
End Sub

Private Sub Idioma_Click()

    If Titulo.Caption = "Valor Standard" Then
    
        Titulo.Caption = "Valor Standard Ingles"
        TituloII.Caption = "Valor Standard Ingles"
        TituloIII.Caption = "Valor Standard Ingles"
        
        Valor1.Visible = False
        valor2.Visible = False
        Valor3.Visible = False
        valor4.Visible = False
        valor5.Visible = False
        valor6.Visible = False
        valor7.Visible = False
        valor8.Visible = False
        valor9.Visible = False
        valor10.Visible = False
        Valor11.Visible = False
        Valor12.Visible = False
        Valor13.Visible = False
        Valor14.Visible = False
        Valor15.Visible = False
        Valor16.Visible = False
        Valor17.Visible = False
        Valor18.Visible = False
        Valor19.Visible = False
        Valor20.Visible = False
        Valor21.Visible = False
        Valor22.Visible = False
        Valor23.Visible = False
        Valor24.Visible = False
        Valor25.Visible = False
        Valor26.Visible = False
        Valor27.Visible = False
        Valor28.Visible = False
        Valor29.Visible = False
        Valor30.Visible = False
        
        IValor1.Visible = True
        IValor2.Visible = True
        IValor3.Visible = True
        IValor4.Visible = True
        IValor5.Visible = True
        IValor6.Visible = True
        IValor7.Visible = True
        IValor8.Visible = True
        IValor9.Visible = True
        IValor10.Visible = True
        IValor11.Visible = True
        IValor12.Visible = True
        IValor13.Visible = True
        IValor14.Visible = True
        IValor15.Visible = True
        IValor16.Visible = True
        IValor17.Visible = True
        IValor18.Visible = True
        IValor19.Visible = True
        IValor20.Visible = True
        IValor21.Visible = True
        IValor22.Visible = True
        IValor23.Visible = True
        IValor24.Visible = True
        IValor25.Visible = True
        IValor26.Visible = True
        IValor27.Visible = True
        IValor28.Visible = True
        IValor29.Visible = True
        IValor30.Visible = True
    
            Else
    
        Titulo.Caption = "Valor Standard"
        TituloII.Caption = "Valor Standard"
        TituloIII.Caption = "Valor Standard"
        
        Valor1.Visible = True
        valor2.Visible = True
        Valor3.Visible = True
        valor4.Visible = True
        valor5.Visible = True
        valor6.Visible = True
        valor7.Visible = True
        valor8.Visible = True
        valor9.Visible = True
        valor10.Visible = True
        Valor11.Visible = True
        Valor12.Visible = True
        Valor13.Visible = True
        Valor14.Visible = True
        Valor15.Visible = True
        Valor16.Visible = True
        Valor17.Visible = True
        Valor18.Visible = True
        Valor19.Visible = True
        Valor20.Visible = True
        Valor21.Visible = True
        Valor22.Visible = True
        Valor23.Visible = True
        Valor24.Visible = True
        Valor25.Visible = True
        Valor26.Visible = True
        Valor27.Visible = True
        Valor28.Visible = True
        Valor29.Visible = True
        Valor30.Visible = True
        
        IValor1.Visible = False
        IValor2.Visible = False
        IValor3.Visible = False
        IValor4.Visible = False
        IValor5.Visible = False
        IValor6.Visible = False
        IValor7.Visible = False
        IValor8.Visible = False
        IValor9.Visible = False
        IValor10.Visible = False
        IValor11.Visible = False
        IValor12.Visible = False
        IValor13.Visible = False
        IValor14.Visible = False
        IValor15.Visible = False
        IValor16.Visible = False
        IValor17.Visible = False
        IValor18.Visible = False
        IValor19.Visible = False
        IValor20.Visible = False
        IValor21.Visible = False
        IValor22.Visible = False
        IValor23.Visible = False
        IValor24.Visible = False
        IValor25.Visible = False
        IValor26.Visible = False
        IValor27.Visible = False
        IValor28.Visible = False
        IValor29.Visible = False
        IValor30.Visible = False
    
    End If

End Sub

Private Sub ImprimeII_Click()
    
    XEmpresa = WEmpresa
    Select Case Val(XEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case 2, 4, 8, 9
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
    End Select
                    
    ZSql = "DELETE CertificadoMp"
    spCertificadoMp = ZSql
    Set rstCertificadoMp = db.OpenRecordset(spCertificadoMp, dbOpenSnapshot, dbSQLPassThrough)
                    
    Erase ZEnsayo
        
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
    
        ZEnsayo(1) = rstEspecificacionesUnifica!Ensayo1
        ZEnsayo(2) = rstEspecificacionesUnifica!Ensayo2
        ZEnsayo(3) = rstEspecificacionesUnifica!Ensayo3
        ZEnsayo(4) = rstEspecificacionesUnifica!Ensayo4
        ZEnsayo(5) = rstEspecificacionesUnifica!Ensayo5
        ZEnsayo(6) = rstEspecificacionesUnifica!Ensayo6
        ZEnsayo(7) = rstEspecificacionesUnifica!Ensayo7
        ZEnsayo(8) = rstEspecificacionesUnifica!Ensayo8
        ZEnsayo(9) = rstEspecificacionesUnifica!Ensayo9
        ZEnsayo(10) = rstEspecificacionesUnifica!Ensayo10
        ZEnsayo(11) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo11), "", rstEspecificacionesUnifica!Ensayo11)
        ZEnsayo(12) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo12), "", rstEspecificacionesUnifica!Ensayo12)
        ZEnsayo(13) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo13), "", rstEspecificacionesUnifica!Ensayo13)
        ZEnsayo(14) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo14), "", rstEspecificacionesUnifica!Ensayo14)
        ZEnsayo(15) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo15), "", rstEspecificacionesUnifica!Ensayo15)
        ZEnsayo(16) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo16), "", rstEspecificacionesUnifica!Ensayo16)
        ZEnsayo(17) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo17), "", rstEspecificacionesUnifica!Ensayo17)
        ZEnsayo(18) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo18), "", rstEspecificacionesUnifica!Ensayo18)
        ZEnsayo(19) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo19), "", rstEspecificacionesUnifica!Ensayo19)
        ZEnsayo(20) = IIf(IsNull(rstEspecificacionesUnifica!Ensayo20), "", rstEspecificacionesUnifica!Ensayo20)
        
        rstEspecificacionesUnifica.Close
                        
    End If
    
    
        
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnificaIII"
    Sql3 = " Where EspecificacionesUnificaIII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaIII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaIII = db.OpenRecordset(spEspecificacionesUnificaIII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaIII.RecordCount > 0 Then
    
        ZEnsayo(21) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo21), "", rstEspecificacionesUnificaIII!Ensayo21)
        ZEnsayo(22) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo22), "", rstEspecificacionesUnificaIII!Ensayo22)
        ZEnsayo(23) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo23), "", rstEspecificacionesUnificaIII!Ensayo23)
        ZEnsayo(24) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo24), "", rstEspecificacionesUnificaIII!Ensayo24)
        ZEnsayo(25) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo25), "", rstEspecificacionesUnificaIII!Ensayo25)
        ZEnsayo(26) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo26), "", rstEspecificacionesUnificaIII!Ensayo26)
        ZEnsayo(27) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo27), "", rstEspecificacionesUnificaIII!Ensayo27)
        ZEnsayo(28) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo28), "", rstEspecificacionesUnificaIII!Ensayo28)
        ZEnsayo(29) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo29), "", rstEspecificacionesUnificaIII!Ensayo29)
        ZEnsayo(30) = IIf(IsNull(rstEspecificacionesUnificaIII!Ensayo30), "", rstEspecificacionesUnificaIII!Ensayo30)
        
        rstEspecificacionesUnificaIII.Close
                        
    End If
    
    
    ZZIValor1 = ""
    ZZIValor2 = ""
    ZZIValor3 = ""
    ZZIValor4 = ""
    ZZIValor5 = ""
    ZZIValor6 = ""
    ZZIValor7 = ""
    ZZIValor8 = ""
    ZZIValor9 = ""
    ZZIValor10 = ""
    ZZIValor11 = ""
    ZZIValor12 = ""
    ZZIValor13 = ""
    ZZIValor14 = ""
    ZZIValor15 = ""
    ZZIValor16 = ""
    ZZIValor17 = ""
    ZZIValor18 = ""
    ZZIValor19 = ""
    ZZIValor20 = ""
    ZZIValor21 = ""
    ZZIValor22 = ""
    ZZIValor23 = ""
    ZZIValor24 = ""
    ZZIValor25 = ""
    ZZIValor26 = ""
    ZZIValor27 = ""
    ZZIValor28 = ""
    ZZIValor29 = ""
    ZZIValor30 = ""
    
    
    Sql1 = "Select EspecificacionesUnificaII.IValor1, EspecificacionesUnificaII.IValor2, EspecificacionesUnificaII.IValor3, EspecificacionesUnificaII.IValor4, EspecificacionesUnificaII.IValor5, EspecificacionesUnificaII.IValor6, EspecificacionesUnificaII.IValor7, EspecificacionesUnificaII.IValor8, EspecificacionesUnificaII.IValor9, EspecificacionesUnificaII.IValor10, "
    Sql2 = "       EspecificacionesUnificaII.IValor11, EspecificacionesUnificaII.IValor12, EspecificacionesUnificaII.IValor13, EspecificacionesUnificaII.IValor14, EspecificacionesUnificaII.IValor15, EspecificacionesUnificaII.IValor16, EspecificacionesUnificaII.IValor17, EspecificacionesUnificaII.IValor18, EspecificacionesUnificaII.IValor19, EspecificacionesUnificaII.IValor20, "
    Sql3 = "       EspecificacionesUnificaII.DescripcionIngles, EspecificacionesUnificaII.cas"
    Sql4 = " FROM EspecificacionesUnificaII"
    Sql5 = " Where EspecificacionesUnificaII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3 + Sql4 + Sql5
    Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaII.RecordCount > 0 Then
        
        ZZDescripcionIngles = IIf(IsNull(rstEspecificacionesUnificaII!DescripcionIngles), "", rstEspecificacionesUnificaII!DescripcionIngles)
        ZZCas = IIf(IsNull(rstEspecificacionesUnificaII!Cas), "", rstEspecificacionesUnificaII!Cas)
       
        ZZDescripcionIngles = Trim(ZZDescripcionIngles)
        ZZCas = Trim(ZZCas)
        
        ZDescriII(1) = IIf(IsNull(rstEspecificacionesUnificaII!IValor1), "", rstEspecificacionesUnificaII!IValor1)
        ZDescriII(2) = IIf(IsNull(rstEspecificacionesUnificaII!IValor2), "", rstEspecificacionesUnificaII!IValor2)
        ZDescriII(3) = IIf(IsNull(rstEspecificacionesUnificaII!IValor3), "", rstEspecificacionesUnificaII!IValor3)
        ZDescriII(4) = IIf(IsNull(rstEspecificacionesUnificaII!IValor4), "", rstEspecificacionesUnificaII!IValor4)
        ZDescriII(5) = IIf(IsNull(rstEspecificacionesUnificaII!IValor5), "", rstEspecificacionesUnificaII!IValor5)
        ZDescriII(6) = IIf(IsNull(rstEspecificacionesUnificaII!IValor6), "", rstEspecificacionesUnificaII!IValor6)
        ZDescriII(7) = IIf(IsNull(rstEspecificacionesUnificaII!IValor7), "", rstEspecificacionesUnificaII!IValor7)
        ZDescriII(8) = IIf(IsNull(rstEspecificacionesUnificaII!IValor8), "", rstEspecificacionesUnificaII!IValor8)
        ZDescriII(9) = IIf(IsNull(rstEspecificacionesUnificaII!IValor9), "", rstEspecificacionesUnificaII!IValor9)
        ZDescriII(10) = IIf(IsNull(rstEspecificacionesUnificaII!IValor10), "", rstEspecificacionesUnificaII!IValor10)
        ZDescriII(11) = IIf(IsNull(rstEspecificacionesUnificaII!IValor11), "", rstEspecificacionesUnificaII!IValor11)
        ZDescriII(12) = IIf(IsNull(rstEspecificacionesUnificaII!IValor12), "", rstEspecificacionesUnificaII!IValor12)
        ZDescriII(13) = IIf(IsNull(rstEspecificacionesUnificaII!IValor13), "", rstEspecificacionesUnificaII!IValor13)
        ZDescriII(14) = IIf(IsNull(rstEspecificacionesUnificaII!IValor14), "", rstEspecificacionesUnificaII!IValor14)
        ZDescriII(15) = IIf(IsNull(rstEspecificacionesUnificaII!IValor15), "", rstEspecificacionesUnificaII!IValor15)
        ZDescriII(16) = IIf(IsNull(rstEspecificacionesUnificaII!IValor16), "", rstEspecificacionesUnificaII!IValor16)
        ZDescriII(17) = IIf(IsNull(rstEspecificacionesUnificaII!IValor17), "", rstEspecificacionesUnificaII!IValor17)
        ZDescriII(18) = IIf(IsNull(rstEspecificacionesUnificaII!IValor18), "", rstEspecificacionesUnificaII!IValor18)
        ZDescriII(19) = IIf(IsNull(rstEspecificacionesUnificaII!IValor19), "", rstEspecificacionesUnificaII!IValor19)
        ZDescriII(20) = IIf(IsNull(rstEspecificacionesUnificaII!IValor20), "", rstEspecificacionesUnificaII!IValor20)
        
        rstEspecificacionesUnificaII.Close
                        
    End If
    
    
    
    Sql1 = "Select EspecificacionesUnificaII.IValor21, EspecificacionesUnificaII.IValor22, EspecificacionesUnificaII.IValor23, EspecificacionesUnificaII.IValor24, EspecificacionesUnificaII.IValor25, EspecificacionesUnificaII.IValor26, EspecificacionesUnificaII.IValor27, EspecificacionesUnificaII.IValor28, EspecificacionesUnificaII.IValor29, EspecificacionesUnificaII.IValor30 "
    Sql2 = " FROM EspecificacionesUnificaII"
    Sql3 = " Where EspecificacionesUnificaII.Producto = " + "'" + Codigo.Text + "'"
    spEspecificacionesUnificaII = Sql1 + Sql2 + Sql3
    Set rstEspecificacionesUnificaII = db.OpenRecordset(spEspecificacionesUnificaII, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnificaII.RecordCount > 0 Then
        
        ZDescriII(21) = IIf(IsNull(rstEspecificacionesUnificaII!IValor21), "", rstEspecificacionesUnificaII!IValor21)
        ZDescriII(22) = IIf(IsNull(rstEspecificacionesUnificaII!IValor22), "", rstEspecificacionesUnificaII!IValor22)
        ZDescriII(23) = IIf(IsNull(rstEspecificacionesUnificaII!IValor23), "", rstEspecificacionesUnificaII!IValor23)
        ZDescriII(24) = IIf(IsNull(rstEspecificacionesUnificaII!IValor24), "", rstEspecificacionesUnificaII!IValor24)
        ZDescriII(25) = IIf(IsNull(rstEspecificacionesUnificaII!IValor25), "", rstEspecificacionesUnificaII!IValor25)
        ZDescriII(26) = IIf(IsNull(rstEspecificacionesUnificaII!IValor26), "", rstEspecificacionesUnificaII!IValor26)
        ZDescriII(27) = IIf(IsNull(rstEspecificacionesUnificaII!IValor27), "", rstEspecificacionesUnificaII!IValor27)
        ZDescriII(28) = IIf(IsNull(rstEspecificacionesUnificaII!IValor28), "", rstEspecificacionesUnificaII!IValor28)
        ZDescriII(29) = IIf(IsNull(rstEspecificacionesUnificaII!IValor29), "", rstEspecificacionesUnificaII!IValor29)
        ZDescriII(30) = IIf(IsNull(rstEspecificacionesUnificaII!IValor30), "", rstEspecificacionesUnificaII!IValor30)
        
        rstEspecificacionesUnificaII.Close
                        
    End If
    
    
    For Cicla = 1 To 30
        ZZDescri = ""
        If Val(ZEnsayo(Cicla)) <> 0 Then
        
            spEnsayo = "ConsultaEnsayos " + "'" + ZEnsayo(Cicla) + "'"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
                ZZDescri = IIf(IsNull(rstEnsayo!DescripcionII), "", rstEnsayo!DescripcionII)
                rstEnsayo.Close
            End If
            
            Auxi1 = Str$(Cicla)
            Call Ceros(Auxi1, 2)
            
            ZClave = Codigo.Text + Auxi1
                
            ZSql = ""
            ZSql = ZSql + "INSERT INTO CertificadoMp ("
            ZSql = ZSql + "Clave ,"
            ZSql = ZSql + "Terminado ,"
            ZSql = ZSql + "Renglon ,"
            ZSql = ZSql + "Descripcion ,"
            ZSql = ZSql + "Examen ,"
            ZSql = ZSql + "Valor ,"
            ZSql = ZSql + "Version ,"
            ZSql = ZSql + "Cas )"
            ZSql = ZSql + "Values ("
            ZSql = ZSql + "'" + ZClave + "',"
            ZSql = ZSql + "'" + Codigo.Text + "',"
            ZSql = ZSql + "'" + Str$(Cicla) + "',"
            ZSql = ZSql + "'" + Trim(ZZDescripcionIngles) + "',"
            ZSql = ZSql + "'" + Trim(ZZDescri) + "',"
            ZSql = ZSql + "'" + Trim(ZDescriII(Cicla)) + "',"
            ZSql = ZSql + "'" + Version.Text + "',"
            ZSql = ZSql + "'" + Trim(ZZCas) + "')"
    
            spCertificadoMp = ZSql
            Set rstCertificadoMp = db.OpenRecordset(spCertificadoMp, dbOpenSnapshot, dbSQLPassThrough)
            
        End If
                
    Next Cicla
            
    lista.WindowTitle = "Certificado de Analisis"
    lista.WindowTop = 0
    lista.WindowLeft = 0
    lista.WindowWidth = Screen.Width
    lista.WindowHeight = Screen.Height

    lista.Destination = 0
    Rem Listado.Destination = 0
            
    If Val(WEmpresa) = 1 Or Val(WEmpresa) = 3 Or Val(WEmpresa) = 5 Or Val(WEmpresa) = 6 Or Val(WEmpresa) = 7 Or Val(WEmpresa) = 10 Or Val(WEmpresa) = 11 Then
        lista.ReportFileName = "CertificadoMp.rpt"
            Else
        lista.ReportFileName = "CertificadoMpPelli.rpt"
    End If
                
    DbConnect = db.Connect
    DSQ = getDatabase(DbConnect)

    lista.SQLQuery = "SELECT CertificadoMp.Terminado, CertificadoMp.Renglon, CertificadoMp.Renglon, CertificadoMp.Descripcion, CertificadoMp.Examen, CertificadoMp.Valor, CertificadoMp.Cas " _
            + "From " _
            + DSQ + ".dbo.CertificadoMp CertificadoMp " _
            + "Where " _
            + "CertificadoMp.Renglon >= 0 AND " _
            + "CertificadoMp.Renglon <= 999999"

    lista.Connect = Connect()
    
    lista.Destination = 0
    Rem Lista.Destination = 0
    
    lista.Action = 1
    
    Call Conecta_Empresa
        
End Sub

Private Sub Listado_Click()
    Desde.Text = "  -   -   "
    Hasta.Text = "  -   -   "
    ImprePantalla.Value = False
    ImpreListado.Value = True
    Frame2.Visible = True
    Desde.SetFocus
End Sub

Private Sub Valor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde1.SetFocus
    End If
End Sub

Private Sub Desde1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde2.SetFocus
    End If
End Sub

Private Sub Desde2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta2.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde3.SetFocus
    End If
End Sub

Private Sub Desde3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta3.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde4.SetFocus
    End If
End Sub

Private Sub Desde4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta4.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde5.SetFocus
    End If
End Sub

Private Sub Desde5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta5.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo6.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde6.SetFocus
    End If
End Sub

Private Sub Desde6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta6.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo7.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde7.SetFocus
    End If
End Sub

Private Sub Desde7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta7.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo8.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde8.SetFocus
    End If
End Sub

Private Sub Desde8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta8.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo9.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde9.SetFocus
    End If
End Sub

Private Sub Desde9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta9.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo10.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde10.SetFocus
    End If
End Sub

Private Sub Desde10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta10.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTab1.Tab = 1
        Ensayo11.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde11.SetFocus
    End If
End Sub

Private Sub Desde11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta11.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo12.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub




Private Sub Valor12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde12.SetFocus
    End If
End Sub

Private Sub Desde12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta12.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo13.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde13.SetFocus
    End If
End Sub

Private Sub Desde13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta13.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo14.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde14.SetFocus
    End If
End Sub

Private Sub Desde14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta14.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo15.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde15.SetFocus
    End If
End Sub

Private Sub Desde15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta15.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo16.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde16.SetFocus
    End If
End Sub

Private Sub Desde16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta16.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo17.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde17.SetFocus
    End If
End Sub

Private Sub Desde17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta17.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo18.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde18.SetFocus
    End If
End Sub

Private Sub Desde18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta18.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo19.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde19.SetFocus
    End If
End Sub

Private Sub Desde19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta19.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo20.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde20.SetFocus
    End If
End Sub

Private Sub Desde20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta20.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTab1.Tab = 2
        Ensayo21.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde21.SetFocus
    End If
End Sub

Private Sub Desde21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta21.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo22.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde22.SetFocus
    End If
End Sub

Private Sub Desde22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta22.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo23.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde23.SetFocus
    End If
End Sub

Private Sub Desde23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta23.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo24.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde24.SetFocus
    End If
End Sub

Private Sub Desde24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta24.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo25.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde25.SetFocus
    End If
End Sub

Private Sub Desde25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta25.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo26.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde26.SetFocus
    End If
End Sub

Private Sub Desde26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta26.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo27.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde27.SetFocus
    End If
End Sub

Private Sub Desde27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta27.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo28.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde28.SetFocus
    End If
End Sub

Private Sub Desde28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta28.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo29.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde29.SetFocus
    End If
End Sub

Private Sub Desde29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta29.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Ensayo30.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Valor30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde30.SetFocus
    End If
End Sub

Private Sub Desde30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta30.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Private Sub Hasta30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SSTab1.Tab = 0
        Ensayo1.SetFocus
    End If
    Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub












Private Sub IValor1_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde1.SetFocus
    End If
End Sub


Private Sub IValor2_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde2.SetFocus
    End If
End Sub


Private Sub IValor3_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde3.SetFocus
    End If
End Sub


Private Sub IValor4_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde4.SetFocus
    End If
End Sub

Private Sub IValor5_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde5.SetFocus
    End If
End Sub


Private Sub IValor6_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde6.SetFocus
    End If
End Sub


Private Sub IValor7_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde7.SetFocus
    End If
End Sub


Private Sub IValor8_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde8.SetFocus
    End If
End Sub


Private Sub IValor9_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde9.SetFocus
    End If
End Sub


Private Sub IValor10_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde10.SetFocus
    End If
End Sub


Private Sub IValor11_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde11.SetFocus
    End If
End Sub





Private Sub IValor12_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde12.SetFocus
    End If
End Sub

Private Sub IValor13_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde13.SetFocus
    End If
End Sub


Private Sub IValor14_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde14.SetFocus
    End If
End Sub


Private Sub IValor15_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde15.SetFocus
    End If
End Sub


Private Sub IValor16_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde16.SetFocus
    End If
End Sub


Private Sub IValor17_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde17.SetFocus
    End If
End Sub


Private Sub IValor18_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde18.SetFocus
    End If
End Sub


Private Sub IValor19_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde19.SetFocus
    End If
End Sub


Private Sub IValor20_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde20.SetFocus
    End If
End Sub


Private Sub IValor21_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde21.SetFocus
    End If
End Sub

Private Sub IValor22_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde22.SetFocus
    End If
End Sub

Private Sub IValor23_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde23.SetFocus
    End If
End Sub

Private Sub IValor24_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde24.SetFocus
    End If
End Sub

Private Sub IValor25_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde25.SetFocus
    End If
End Sub

Private Sub IValor26_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde26.SetFocus
    End If
End Sub

Private Sub IValor27_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde27.SetFocus
    End If
End Sub

Private Sub IValor28_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde28.SetFocus
    End If
End Sub

Private Sub IValor29_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde29.SetFocus
    End If
End Sub

Private Sub IValor30_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde30.SetFocus
    End If
End Sub






























Sub Codigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Codigo.Text <> "" Then
            Codigo.Text = UCase(Codigo.Text)
            
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
                
                
            Sql1 = "Select EspecificacionesUnifica.Producto"
            Sql2 = " FROM EspecificacionesUnifica"
            Sql3 = " Where EspecificacionesUnifica.Producto = " + "'" + Codigo.Text + "'"
            spEspecificacionesUnifica = Sql1 + Sql2 + Sql3
            Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
            If rstEspecificacionesUnifica.RecordCount > 0 Then
                rstEspecificacionesUnifica.Close
                Call Conecta_Empresa
                Call Imprime_Datos
                    Else
                Call Conecta_Empresa
                WCodigo = Codigo.Text
                CmdLimpiar_Click
                Codigo.Text = WCodigo
            End If
            
            spArticulo = "ConsultaArticulo " + "'" + Codigo.Text + "'"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
                Descriprod.Caption = rstArticulo!Descripcion
                rstArticulo.Close
                    Else
                Codigo.SetFocus
                Exit Sub
            End If
            
        End If
        Ensayo1.SetFocus
    End If
    Rem Call NumbersOnly(Screen.ActiveControl, KeyAscii)
End Sub

Sub Desde_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Desde.Text = UCase(Desde.Text)
        Hasta.SetFocus
    End If
End Sub

Sub Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Hasta.Text = UCase(Hasta.Text)
        Desde.SetFocus
    End If
End Sub

Private Sub Consulta_Click()
    Opcion.Clear
    
    Opcion.AddItem "Codigos"
    Opcion.AddItem "Ensayos"
    
    Opcion.Visible = True
End Sub

Private Sub Opcion_Click()
    Opcion.Visible = False
    Dim IngresaItem As String

    pantalla.Clear
    WIndice.Clear

    XIndice = Opcion.ListIndex
    
    Select Case XIndice
        Case 0
            spArticulo = "ListaArticuloConsulta"
            Set rstArticulo = db.OpenRecordset(spArticulo, dbOpenSnapshot, dbSQLPassThrough)
            If rstArticulo.RecordCount > 0 Then
            
            With rstArticulo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = rstArticulo!Codigo + " " + rstArticulo!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstArticulo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstArticulo.Close
            
            End If
            
        Case 1
            XEmpresa = WEmpresa
            Select Case Val(WEmpresa)
                Case 1, 3, 5, 6, 7, 10, 11
                    WEmpresa = "0003"
                    txtOdbc = "Empresa03"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
                Case Else
                    WEmpresa = "0004"
                    txtOdbc = "Empresa04"
                    strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
                    Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
            End Select
            
            spEnsayo = "ListaEnsayos"
            Set rstEnsayo = db.OpenRecordset(spEnsayo, dbOpenSnapshot, dbSQLPassThrough)
            If rstEnsayo.RecordCount > 0 Then
            
            With rstEnsayo
                .MoveFirst
                Do
                    If .EOF = False Then
                        IngresaItem = Str$(rstEnsayo!Codigo) + " " + rstEnsayo!Descripcion
                        pantalla.AddItem IngresaItem
                        IngresaItem = rstEnsayo!Codigo
                        WIndice.AddItem IngresaItem
                        .MoveNext
                            Else
                        Exit Do
                    End If
                Loop
            End With
            rstEnsayo.Close
            
            End If
            
            Call Conecta_Empresa
            
        Case Else
    End Select
            
    pantalla.Visible = True

End Sub

Private Sub pantalla_Click()
    pantalla.Visible = False
    Select Case XIndice
        Case 0
            Indice = pantalla.ListIndex
            Codigo.Text = WIndice.List(Indice)
            Call Codigo_KeyPress(13)
            
        Case 1
            Entra$ = "S"
            If Val(Ensayo1.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo1.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo1_Keypress(13)
            End If
            If Val(Ensayo2.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo2.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo2_Keypress(13)
            End If
            If Val(Ensayo3.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo3.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo3_Keypress(13)
            End If
            If Val(Ensayo4.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo4.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo4_Keypress(13)
            End If
            If Val(Ensayo5.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo5.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo5_Keypress(13)
            End If
            If Val(Ensayo6.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo6.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo6_Keypress(13)
            End If
            If Val(Ensayo7.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo7.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo7_Keypress(13)
            End If
            If Val(Ensayo8.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo8.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo8_Keypress(13)
            End If
            If Val(Ensayo9.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo9.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo9_Keypress(13)
            End If
            If Val(Ensayo10.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo10.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo10_Keypress(13)
            End If
            If Val(Ensayo11.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo11.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo11_Keypress(13)
            End If
            If Val(Ensayo12.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo12.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo12_Keypress(13)
            End If
            If Val(Ensayo13.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo13.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo13_Keypress(13)
            End If
            If Val(Ensayo14.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo14.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo14_Keypress(13)
            End If
            If Val(Ensayo15.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo15.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo15_Keypress(13)
            End If
            If Val(Ensayo16.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo16.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo16_Keypress(13)
            End If
            If Val(Ensayo17.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo17.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo17_Keypress(13)
            End If
            If Val(Ensayo18.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo18.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo18_Keypress(13)
            End If
            If Val(Ensayo19.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo19.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo19_Keypress(13)
            End If
            If Val(Ensayo20.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo20.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo20_Keypress(13)
            End If
            If Val(Ensayo21.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo21.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo21_Keypress(13)
            End If
            If Val(Ensayo22.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo22.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo22_Keypress(13)
            End If
            If Val(Ensayo23.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo23.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo23_Keypress(13)
            End If
            If Val(Ensayo24.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo24.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo24_Keypress(13)
            End If
            If Val(Ensayo25.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo25.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo25_Keypress(13)
            End If
            If Val(Ensayo26.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo26.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo26_Keypress(13)
            End If
            If Val(Ensayo27.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo27.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo27_Keypress(13)
            End If
            If Val(Ensayo28.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo28.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo28_Keypress(13)
            End If
            If Val(Ensayo29.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo29.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo29_Keypress(13)
            End If
            If Val(Ensayo30.Text) = 0 And Entra$ = "S" Then
                Indice = pantalla.ListIndex
                Ensayo30.Text = Val(WIndice.List(Indice))
                Entra$ = "N"
                Call Ensayo30_Keypress(13)
            End If
        Case Else
    End Select
    
End Sub

Private Sub Anterior_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select

    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto < " + "'" + Codigo.Text + "'"
    Sql4 = " Order by EspecificacionesUnifica.Producto"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        With rstEspecificacionesUnifica
            .MoveLast
            Codigo.Text = rstEspecificacionesUnifica!Producto
        End With
        rstEspecificacionesUnifica.Close
            Else
        m$ = "No exsite registro Anterior"
        A% = MsgBox(m$, 0, "Ingreso de Especificaciones de Materias Primas")
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Codigo.SetFocus
    
End Sub

Private Sub Primer_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Sql1 = "Select Min(Producto) as [ProductoMenor]"
    Sql2 = " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = Sql1 + Sql2
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        rstEspecificacionesUnifica.MoveFirst
        Codigo.Text = rstEspecificacionesUnifica!ProductoMenor
        rstEspecificacionesUnifica.Close
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Codigo.SetFocus
    
 End Sub

Private Sub Ultimo_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Sql1 = "Select Max(Producto) as [ProductoMayor]"
    Sql2 = " FROM EspecificacionesUnifica"
    spEspecificacionesUnifica = Sql1 + Sql2
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        rstEspecificacionesUnifica.MoveLast
        Codigo.Text = rstEspecificacionesUnifica!ProductoMayor
        rstEspecificacionesUnifica.Close
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Codigo.SetFocus

 End Sub

Private Sub Siguiente_Click()

    XEmpresa = WEmpresa
    Select Case Val(WEmpresa)
        Case 1, 3, 5, 6, 7, 10, 11
            WEmpresa = "0003"
            txtOdbc = "Empresa03"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
        Case Else
            WEmpresa = "0004"
            txtOdbc = "Empresa04"
            strConnect = "odbc;dsn=" & txtOdbc & ";uid=" & txtUserName & ";pwd=" & txtPassword & ";app=" & gAplicacion
            Set db = DBEngine.Workspaces(0).OpenDatabase("", False, False, strConnect)
    End Select
    
    Sql1 = "Select *"
    Sql2 = " FROM EspecificacionesUnifica"
    Sql3 = " Where EspecificacionesUnifica.Producto > " + "'" + Codigo.Text + "'"
    Sql4 = " Order by EspecificacionesUnifica.Producto"
    spEspecificacionesUnifica = Sql1 + Sql2 + Sql3 + Sql4
    Set rstEspecificacionesUnifica = db.OpenRecordset(spEspecificacionesUnifica, dbOpenSnapshot, dbSQLPassThrough)
    If rstEspecificacionesUnifica.RecordCount > 0 Then
        With rstEspecificacionesUnifica
            .MoveFirst
            Codigo.Text = rstEspecificacionesUnifica!Producto
        End With
        rstEspecificacionesUnifica.Close
            Else
        m$ = "No exsite registro Posterior"
        A% = MsgBox(m$, 0, "Ingreso de Especificaciones de Materias Primas")
    End If
    
    Call Conecta_Empresa
    
    Call Imprime_Datos
    Codigo.SetFocus
    
End Sub

Sub Ingresa_clave()
    WClave.Text = ""
    XClave.Visible = True
    WClave.SetFocus
End Sub

Private Sub CancelaGraba_Click()
    XClave.Visible = False
End Sub

Private Sub WClave_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
        WGraba = "N"
        ZGRABAII = ""
        
        ZSql = ""
        ZSql = ZSql + "Select *"
        ZSql = ZSql + " FROM Operador"
        ZSql = ZSql + " Where Operador.Clave = " + "'" + WClave.Text + "'"
        spOperador = ZSql
        Set rstOperador = db.OpenRecordset(spOperador, dbOpenSnapshot, dbSQLPassThrough)
        If rstOperador.RecordCount > 0 Then
            ZOperador = rstOperador!Operador
            ZGRABAII = IIf(IsNull(rstOperador!GrabaII), "", rstOperador!GrabaII)
            rstOperador.Close
        End If
        
        If ZGRABAII = "S" Then
            WGraba = "S"
            XClave.Visible = False
            Select Case ZZProceso
                Case 0
                    Call cmdAdd_Click
                Case Else
                    Call GrabaII_Click
            End Select
                Else
            m$ = "Clave de Grabacion Invalida"
            A% = MsgBox(m$, 0, "Especificaciones de Materia Prima")
            WClave.SetFocus
        End If
        
    End If
End Sub

Private Sub DescripcionIngles_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Cas.SetFocus
    End If
    If KeyAscii = 27 Then
        DescripcionIngles.Text = ""
    End If
End Sub

Private Sub Cas_Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        DescripcionIngles.SetFocus
    End If
    If KeyAscii = 27 Then
        Cas.Text = ""
    End If
End Sub

