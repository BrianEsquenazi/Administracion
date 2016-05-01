VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form DemoDesarrolloII 
   Caption         =   "Ingreso de Orden de Trabajo"
   ClientHeight    =   8340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   11625
   Begin VB.TextBox Text42 
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
      Left            =   5400
      MaxLength       =   6
      TabIndex        =   35
      Text            =   " 20"
      Top             =   600
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   11668
      _Version        =   327680
      Tabs            =   6
      Tab             =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso de la Prueba"
      TabPicture(0)   =   "DemoDesarrolloII.frx":0000
      Tab(0).ControlCount=   48
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label18"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Text5"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Text6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Text7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text23"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text24"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text25"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text43"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text44"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text45"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text46"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text47"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text48"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text49"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text50"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text51"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text52"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text53"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text54"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text55"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text56"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text57"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text58"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text59"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text60"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text61"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text62"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text63"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text64"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text65"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text66"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text67"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text68"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text69"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text70"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text71"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text72"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text73"
      Tab(0).Control(47).Enabled=   0   'False
      TabCaption(1)   =   "Procedimiento"
      TabPicture(1)   =   "DemoDesarrolloII.frx":001C
      Tab(1).ControlCount=   63
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label23"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label24"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label25"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Text11"
      Tab(1).Control(7).Enabled=   -1  'True
      Tab(1).Control(8)=   "Text10"
      Tab(1).Control(8).Enabled=   -1  'True
      Tab(1).Control(9)=   "Text12"
      Tab(1).Control(9).Enabled=   -1  'True
      Tab(1).Control(10)=   "Text13"
      Tab(1).Control(10).Enabled=   -1  'True
      Tab(1).Control(11)=   "Text14"
      Tab(1).Control(11).Enabled=   -1  'True
      Tab(1).Control(12)=   "Text15"
      Tab(1).Control(12).Enabled=   -1  'True
      Tab(1).Control(13)=   "Text16"
      Tab(1).Control(13).Enabled=   -1  'True
      Tab(1).Control(14)=   "Text17"
      Tab(1).Control(14).Enabled=   -1  'True
      Tab(1).Control(15)=   "Text74"
      Tab(1).Control(15).Enabled=   -1  'True
      Tab(1).Control(16)=   "Text75"
      Tab(1).Control(16).Enabled=   -1  'True
      Tab(1).Control(17)=   "Text76"
      Tab(1).Control(17).Enabled=   -1  'True
      Tab(1).Control(18)=   "Text77"
      Tab(1).Control(18).Enabled=   -1  'True
      Tab(1).Control(19)=   "Text78"
      Tab(1).Control(19).Enabled=   -1  'True
      Tab(1).Control(20)=   "Text79"
      Tab(1).Control(20).Enabled=   -1  'True
      Tab(1).Control(21)=   "Text80"
      Tab(1).Control(21).Enabled=   -1  'True
      Tab(1).Control(22)=   "Text81"
      Tab(1).Control(22).Enabled=   -1  'True
      Tab(1).Control(23)=   "Text82"
      Tab(1).Control(23).Enabled=   -1  'True
      Tab(1).Control(24)=   "Text83"
      Tab(1).Control(24).Enabled=   -1  'True
      Tab(1).Control(25)=   "Text84"
      Tab(1).Control(25).Enabled=   -1  'True
      Tab(1).Control(26)=   "Text85"
      Tab(1).Control(26).Enabled=   -1  'True
      Tab(1).Control(27)=   "Text86"
      Tab(1).Control(27).Enabled=   -1  'True
      Tab(1).Control(28)=   "Text87"
      Tab(1).Control(28).Enabled=   -1  'True
      Tab(1).Control(29)=   "Text88"
      Tab(1).Control(29).Enabled=   -1  'True
      Tab(1).Control(30)=   "Text89"
      Tab(1).Control(30).Enabled=   -1  'True
      Tab(1).Control(31)=   "Text90"
      Tab(1).Control(31).Enabled=   -1  'True
      Tab(1).Control(32)=   "Text91"
      Tab(1).Control(32).Enabled=   -1  'True
      Tab(1).Control(33)=   "Text92"
      Tab(1).Control(33).Enabled=   -1  'True
      Tab(1).Control(34)=   "Text93"
      Tab(1).Control(34).Enabled=   -1  'True
      Tab(1).Control(35)=   "Text94"
      Tab(1).Control(35).Enabled=   -1  'True
      Tab(1).Control(36)=   "Text95"
      Tab(1).Control(36).Enabled=   -1  'True
      Tab(1).Control(37)=   "Text96"
      Tab(1).Control(37).Enabled=   -1  'True
      Tab(1).Control(38)=   "Text97"
      Tab(1).Control(38).Enabled=   -1  'True
      Tab(1).Control(39)=   "Text98"
      Tab(1).Control(39).Enabled=   -1  'True
      Tab(1).Control(40)=   "Text99"
      Tab(1).Control(40).Enabled=   -1  'True
      Tab(1).Control(41)=   "Text100"
      Tab(1).Control(41).Enabled=   -1  'True
      Tab(1).Control(42)=   "Text154"
      Tab(1).Control(42).Enabled=   -1  'True
      Tab(1).Control(43)=   "Text155"
      Tab(1).Control(43).Enabled=   -1  'True
      Tab(1).Control(44)=   "Text156"
      Tab(1).Control(44).Enabled=   -1  'True
      Tab(1).Control(45)=   "Text157"
      Tab(1).Control(45).Enabled=   -1  'True
      Tab(1).Control(46)=   "Text158"
      Tab(1).Control(46).Enabled=   -1  'True
      Tab(1).Control(47)=   "Text159"
      Tab(1).Control(47).Enabled=   -1  'True
      Tab(1).Control(48)=   "Text160"
      Tab(1).Control(48).Enabled=   -1  'True
      Tab(1).Control(49)=   "Text161"
      Tab(1).Control(49).Enabled=   -1  'True
      Tab(1).Control(50)=   "Text162"
      Tab(1).Control(50).Enabled=   -1  'True
      Tab(1).Control(51)=   "Text163"
      Tab(1).Control(51).Enabled=   -1  'True
      Tab(1).Control(52)=   "Text164"
      Tab(1).Control(52).Enabled=   -1  'True
      Tab(1).Control(53)=   "Text165"
      Tab(1).Control(53).Enabled=   -1  'True
      Tab(1).Control(54)=   "Text166"
      Tab(1).Control(54).Enabled=   -1  'True
      Tab(1).Control(55)=   "Text167"
      Tab(1).Control(55).Enabled=   -1  'True
      Tab(1).Control(56)=   "Text168"
      Tab(1).Control(56).Enabled=   -1  'True
      Tab(1).Control(57)=   "Text169"
      Tab(1).Control(57).Enabled=   -1  'True
      Tab(1).Control(58)=   "Text170"
      Tab(1).Control(58).Enabled=   -1  'True
      Tab(1).Control(59)=   "Text171"
      Tab(1).Control(59).Enabled=   -1  'True
      Tab(1).Control(60)=   "Text172"
      Tab(1).Control(60).Enabled=   -1  'True
      Tab(1).Control(61)=   "Text173"
      Tab(1).Control(61).Enabled=   -1  'True
      Tab(1).Control(62)=   "Text174"
      Tab(1).Control(62).Enabled=   -1  'True
      TabCaption(2)   =   "Especificaciones"
      TabPicture(2)   =   "DemoDesarrolloII.frx":0038
      Tab(2).ControlCount=   26
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label17"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label19"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label20"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label21"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label22"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text18"
      Tab(2).Control(5).Enabled=   -1  'True
      Tab(2).Control(6)=   "Text26"
      Tab(2).Control(6).Enabled=   -1  'True
      Tab(2).Control(7)=   "Text19"
      Tab(2).Control(7).Enabled=   -1  'True
      Tab(2).Control(8)=   "Text20"
      Tab(2).Control(8).Enabled=   -1  'True
      Tab(2).Control(9)=   "Text21"
      Tab(2).Control(9).Enabled=   -1  'True
      Tab(2).Control(10)=   "Text22"
      Tab(2).Control(10).Enabled=   -1  'True
      Tab(2).Control(11)=   "Text27"
      Tab(2).Control(11).Enabled=   -1  'True
      Tab(2).Control(12)=   "Text28"
      Tab(2).Control(12).Enabled=   -1  'True
      Tab(2).Control(13)=   "Text29"
      Tab(2).Control(13).Enabled=   -1  'True
      Tab(2).Control(14)=   "Text30"
      Tab(2).Control(14).Enabled=   -1  'True
      Tab(2).Control(15)=   "Text31"
      Tab(2).Control(15).Enabled=   -1  'True
      Tab(2).Control(16)=   "Text32"
      Tab(2).Control(16).Enabled=   -1  'True
      Tab(2).Control(17)=   "Text33"
      Tab(2).Control(17).Enabled=   -1  'True
      Tab(2).Control(18)=   "Text34"
      Tab(2).Control(18).Enabled=   -1  'True
      Tab(2).Control(19)=   "Text35"
      Tab(2).Control(19).Enabled=   -1  'True
      Tab(2).Control(20)=   "Text36"
      Tab(2).Control(20).Enabled=   -1  'True
      Tab(2).Control(21)=   "Text37"
      Tab(2).Control(21).Enabled=   -1  'True
      Tab(2).Control(22)=   "Text38"
      Tab(2).Control(22).Enabled=   -1  'True
      Tab(2).Control(23)=   "Text39"
      Tab(2).Control(23).Enabled=   -1  'True
      Tab(2).Control(24)=   "Text40"
      Tab(2).Control(24).Enabled=   -1  'True
      Tab(2).Control(25)=   "Text41"
      Tab(2).Control(25).Enabled=   -1  'True
      TabCaption(3)   =   "Controles"
      TabPicture(3)   =   "DemoDesarrolloII.frx":0054
      Tab(3).ControlCount=   11
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label34"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Text175"
      Tab(3).Control(1).Enabled=   -1  'True
      Tab(3).Control(2)=   "Text176"
      Tab(3).Control(2).Enabled=   -1  'True
      Tab(3).Control(3)=   "Texfdgh"
      Tab(3).Control(3).Enabled=   -1  'True
      Tab(3).Control(4)=   "Text177"
      Tab(3).Control(4).Enabled=   -1  'True
      Tab(3).Control(5)=   "Text178"
      Tab(3).Control(5).Enabled=   -1  'True
      Tab(3).Control(6)=   "Text179"
      Tab(3).Control(6).Enabled=   -1  'True
      Tab(3).Control(7)=   "Text180"
      Tab(3).Control(7).Enabled=   -1  'True
      Tab(3).Control(8)=   "Text181"
      Tab(3).Control(8).Enabled=   -1  'True
      Tab(3).Control(9)=   "Text182"
      Tab(3).Control(9).Enabled=   -1  'True
      Tab(3).Control(10)=   "Text183"
      Tab(3).Control(10).Enabled=   -1  'True
      TabCaption(4)   =   "Resumen /  Nota"
      TabPicture(4)   =   "DemoDesarrolloII.frx":0070
      Tab(4).ControlCount=   15
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label26"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label27"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label35"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Text101"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Text102"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Text184"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Text185"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Text186"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "Text187"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "Text188"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).Control(10)=   "Text189"
      Tab(4).Control(10).Enabled=   0   'False
      Tab(4).Control(11)=   "Text190"
      Tab(4).Control(11).Enabled=   0   'False
      Tab(4).Control(12)=   "Text191"
      Tab(4).Control(12).Enabled=   0   'False
      Tab(4).Control(13)=   "Text192"
      Tab(4).Control(13).Enabled=   0   'False
      Tab(4).Control(14)=   "Text193"
      Tab(4).Control(14).Enabled=   0   'False
      TabCaption(5)   =   "Costo"
      TabPicture(5)   =   "DemoDesarrolloII.frx":008C
      Tab(5).ControlCount=   60
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label14"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "Label15"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "Label16"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label28"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "Label29"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "Label30"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "Label31"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "Label32"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "Label33"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "Text103"
      Tab(5).Control(9).Enabled=   -1  'True
      Tab(5).Control(10)=   "Text104"
      Tab(5).Control(10).Enabled=   -1  'True
      Tab(5).Control(11)=   "Text105"
      Tab(5).Control(11).Enabled=   -1  'True
      Tab(5).Control(12)=   "Text106"
      Tab(5).Control(12).Enabled=   -1  'True
      Tab(5).Control(13)=   "Text107"
      Tab(5).Control(13).Enabled=   -1  'True
      Tab(5).Control(14)=   "Text108"
      Tab(5).Control(14).Enabled=   -1  'True
      Tab(5).Control(15)=   "Text109"
      Tab(5).Control(15).Enabled=   -1  'True
      Tab(5).Control(16)=   "Text110"
      Tab(5).Control(16).Enabled=   -1  'True
      Tab(5).Control(17)=   "Text111"
      Tab(5).Control(17).Enabled=   -1  'True
      Tab(5).Control(18)=   "Text112"
      Tab(5).Control(18).Enabled=   -1  'True
      Tab(5).Control(19)=   "Text113"
      Tab(5).Control(19).Enabled=   -1  'True
      Tab(5).Control(20)=   "Text114"
      Tab(5).Control(20).Enabled=   -1  'True
      Tab(5).Control(21)=   "Text115"
      Tab(5).Control(21).Enabled=   -1  'True
      Tab(5).Control(22)=   "Text116"
      Tab(5).Control(22).Enabled=   -1  'True
      Tab(5).Control(23)=   "Text117"
      Tab(5).Control(23).Enabled=   -1  'True
      Tab(5).Control(24)=   "Text118"
      Tab(5).Control(24).Enabled=   -1  'True
      Tab(5).Control(25)=   "Text119"
      Tab(5).Control(25).Enabled=   -1  'True
      Tab(5).Control(26)=   "Text120"
      Tab(5).Control(26).Enabled=   -1  'True
      Tab(5).Control(27)=   "Text121"
      Tab(5).Control(27).Enabled=   -1  'True
      Tab(5).Control(28)=   "Text122"
      Tab(5).Control(28).Enabled=   -1  'True
      Tab(5).Control(29)=   "Text123"
      Tab(5).Control(29).Enabled=   -1  'True
      Tab(5).Control(30)=   "Text124"
      Tab(5).Control(30).Enabled=   -1  'True
      Tab(5).Control(31)=   "Text125"
      Tab(5).Control(31).Enabled=   -1  'True
      Tab(5).Control(32)=   "Text126"
      Tab(5).Control(32).Enabled=   -1  'True
      Tab(5).Control(33)=   "Text127"
      Tab(5).Control(33).Enabled=   -1  'True
      Tab(5).Control(34)=   "Text128"
      Tab(5).Control(34).Enabled=   -1  'True
      Tab(5).Control(35)=   "Text129"
      Tab(5).Control(35).Enabled=   -1  'True
      Tab(5).Control(36)=   "Text130"
      Tab(5).Control(36).Enabled=   -1  'True
      Tab(5).Control(37)=   "Text131"
      Tab(5).Control(37).Enabled=   -1  'True
      Tab(5).Control(38)=   "Text132"
      Tab(5).Control(38).Enabled=   -1  'True
      Tab(5).Control(39)=   "Text133"
      Tab(5).Control(39).Enabled=   -1  'True
      Tab(5).Control(40)=   "Text134"
      Tab(5).Control(40).Enabled=   -1  'True
      Tab(5).Control(41)=   "Text135"
      Tab(5).Control(41).Enabled=   -1  'True
      Tab(5).Control(42)=   "Text136"
      Tab(5).Control(42).Enabled=   -1  'True
      Tab(5).Control(43)=   "Text137"
      Tab(5).Control(43).Enabled=   -1  'True
      Tab(5).Control(44)=   "Text138"
      Tab(5).Control(44).Enabled=   -1  'True
      Tab(5).Control(45)=   "Text139"
      Tab(5).Control(45).Enabled=   -1  'True
      Tab(5).Control(46)=   "Text140"
      Tab(5).Control(46).Enabled=   -1  'True
      Tab(5).Control(47)=   "Text141"
      Tab(5).Control(47).Enabled=   -1  'True
      Tab(5).Control(48)=   "Text142"
      Tab(5).Control(48).Enabled=   -1  'True
      Tab(5).Control(49)=   "Text143"
      Tab(5).Control(49).Enabled=   -1  'True
      Tab(5).Control(50)=   "Text144"
      Tab(5).Control(50).Enabled=   -1  'True
      Tab(5).Control(51)=   "Text145"
      Tab(5).Control(51).Enabled=   -1  'True
      Tab(5).Control(52)=   "Text146"
      Tab(5).Control(52).Enabled=   -1  'True
      Tab(5).Control(53)=   "Text147"
      Tab(5).Control(53).Enabled=   -1  'True
      Tab(5).Control(54)=   "Text148"
      Tab(5).Control(54).Enabled=   -1  'True
      Tab(5).Control(55)=   "Text149"
      Tab(5).Control(55).Enabled=   -1  'True
      Tab(5).Control(56)=   "Text150"
      Tab(5).Control(56).Enabled=   -1  'True
      Tab(5).Control(57)=   "Text151"
      Tab(5).Control(57).Enabled=   -1  'True
      Tab(5).Control(58)=   "Text152"
      Tab(5).Control(58).Enabled=   -1  'True
      Tab(5).Control(59)=   "Text153"
      Tab(5).Control(59).Enabled=   -1  'True
      Begin VB.TextBox Text193 
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
         MaxLength       =   50
         TabIndex        =   231
         Text            =   " "
         Top             =   5040
         Width           =   5895
      End
      Begin VB.TextBox Text192 
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
         MaxLength       =   50
         TabIndex        =   230
         Text            =   " "
         Top             =   4680
         Width           =   5895
      End
      Begin VB.TextBox Text191 
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
         MaxLength       =   50
         TabIndex        =   229
         Text            =   " "
         Top             =   3960
         Width           =   5895
      End
      Begin VB.TextBox Text190 
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
         MaxLength       =   50
         TabIndex        =   228
         Text            =   " "
         Top             =   3600
         Width           =   5895
      End
      Begin VB.TextBox Text189 
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
         MaxLength       =   50
         TabIndex        =   227
         Text            =   " "
         Top             =   4320
         Width           =   5895
      End
      Begin VB.TextBox Text188 
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
         MaxLength       =   50
         TabIndex        =   226
         Text            =   " "
         Top             =   3240
         Width           =   5895
      End
      Begin VB.TextBox Text187 
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
         MaxLength       =   50
         TabIndex        =   225
         Text            =   " "
         Top             =   2880
         Width           =   5895
      End
      Begin VB.TextBox Text186 
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
         MaxLength       =   50
         TabIndex        =   223
         Text            =   " "
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox Text185 
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
         MaxLength       =   50
         TabIndex        =   222
         Text            =   " "
         Top             =   1800
         Width           =   5895
      End
      Begin VB.TextBox Text184 
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
         MaxLength       =   50
         TabIndex        =   221
         Text            =   " "
         Top             =   2520
         Width           =   5895
      End
      Begin VB.TextBox Text183 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   220
         Text            =   " "
         Top             =   4200
         Width           =   5895
      End
      Begin VB.TextBox Text182 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   219
         Text            =   " "
         Top             =   3840
         Width           =   5895
      End
      Begin VB.TextBox Text181 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   218
         Text            =   " "
         Top             =   3120
         Width           =   5895
      End
      Begin VB.TextBox Text180 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   217
         Text            =   " "
         Top             =   2760
         Width           =   5895
      End
      Begin VB.TextBox Text179 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   216
         Text            =   " "
         Top             =   3480
         Width           =   5895
      End
      Begin VB.TextBox Text178 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   215
         Text            =   " "
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox Text177 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   214
         Text            =   " "
         Top             =   2040
         Width           =   5895
      End
      Begin VB.TextBox Texfdgh 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   213
         Text            =   " "
         Top             =   1680
         Width           =   5895
      End
      Begin VB.TextBox Text176 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   211
         Text            =   " "
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox Text175 
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
         Left            =   -71520
         MaxLength       =   50
         TabIndex        =   210
         Text            =   " "
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox Text174 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   209
         Text            =   " "
         Top             =   4095
         Width           =   975
      End
      Begin VB.TextBox Text173 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   208
         Text            =   " "
         Top             =   4095
         Width           =   975
      End
      Begin VB.TextBox Text172 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   207
         Text            =   " "
         Top             =   4095
         Width           =   855
      End
      Begin VB.TextBox Text171 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   206
         Text            =   " "
         Top             =   3705
         Width           =   975
      End
      Begin VB.TextBox Text170 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   205
         Text            =   " "
         Top             =   3705
         Width           =   975
      End
      Begin VB.TextBox Text169 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   204
         Text            =   " "
         Top             =   3705
         Width           =   855
      End
      Begin VB.TextBox Text168 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   203
         Text            =   " "
         Top             =   3345
         Width           =   975
      End
      Begin VB.TextBox Text167 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   202
         Text            =   " "
         Top             =   3345
         Width           =   975
      End
      Begin VB.TextBox Text166 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   201
         Text            =   " "
         Top             =   3345
         Width           =   855
      End
      Begin VB.TextBox Text165 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   200
         Text            =   " "
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text164 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   199
         Text            =   " "
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox Text163 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   198
         Text            =   " "
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Text162 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   197
         Text            =   " "
         Top             =   2600
         Width           =   975
      End
      Begin VB.TextBox Text161 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   196
         Text            =   " "
         Top             =   2600
         Width           =   975
      End
      Begin VB.TextBox Text160 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   195
         Text            =   " "
         Top             =   2600
         Width           =   855
      End
      Begin VB.TextBox Text159 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   194
         Text            =   " "
         Top             =   2200
         Width           =   975
      End
      Begin VB.TextBox Text158 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   193
         Text            =   " "
         Top             =   2200
         Width           =   975
      End
      Begin VB.TextBox Text157 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   192
         Text            =   " "
         Top             =   2200
         Width           =   855
      End
      Begin VB.TextBox Text156 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   191
         Text            =   " "
         Top             =   1850
         Width           =   975
      End
      Begin VB.TextBox Text155 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   190
         Text            =   " "
         Top             =   1850
         Width           =   975
      End
      Begin VB.TextBox Text154 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   189
         Text            =   " "
         Top             =   1850
         Width           =   855
      End
      Begin VB.TextBox Text153 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   186
         Text            =   " 4.34"
         Top             =   4680
         Width           =   1095
      End
      Begin VB.TextBox Text152 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   185
         Text            =   "86.88"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox Text151 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   184
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text150 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   183
         Text            =   " "
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text149 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   182
         Text            =   " "
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text148 
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
         Left            =   -65040
         MaxLength       =   6
         TabIndex        =   181
         Text            =   " "
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text147 
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
         Left            =   -65040
         MaxLength       =   20
         TabIndex        =   180
         Text            =   "29.48"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text146 
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
         Left            =   -65040
         MaxLength       =   20
         TabIndex        =   179
         Text            =   "33.31"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text145 
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
         Left            =   -65040
         MaxLength       =   20
         TabIndex        =   177
         Text            =   "24.09"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text144 
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
         Left            =   -66240
         MaxLength       =   6
         TabIndex        =   176
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text143 
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
         MaxLength       =   6
         TabIndex        =   175
         Text            =   " "
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox Text142 
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
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   174
         Text            =   " "
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox Text141 
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
         Left            =   -72720
         MaxLength       =   6
         TabIndex        =   173
         Text            =   " "
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text140 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   172
         Text            =   " "
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox Text139 
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
         MaxLength       =   6
         TabIndex        =   171
         Text            =   " "
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox Text138 
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
         Left            =   -66240
         MaxLength       =   6
         TabIndex        =   170
         Text            =   " "
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text137 
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
         MaxLength       =   6
         TabIndex        =   169
         Text            =   " "
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox Text136 
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
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   168
         Text            =   " "
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox Text135 
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
         Left            =   -72720
         MaxLength       =   6
         TabIndex        =   167
         Text            =   " "
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox Text134 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   166
         Text            =   " "
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox Text133 
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
         MaxLength       =   6
         TabIndex        =   165
         Text            =   " "
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox Text132 
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
         Left            =   -66240
         MaxLength       =   6
         TabIndex        =   164
         Text            =   " "
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text131 
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
         MaxLength       =   6
         TabIndex        =   163
         Text            =   " "
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox Text130 
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
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   162
         Text            =   " "
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox Text129 
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
         Left            =   -72720
         MaxLength       =   6
         TabIndex        =   161
         Text            =   " "
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text128 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   160
         Text            =   " "
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox Text127 
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
         MaxLength       =   6
         TabIndex        =   159
         Text            =   " "
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox Text126 
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
         Left            =   -66240
         MaxLength       =   6
         TabIndex        =   158
         Text            =   " "
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text125 
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
         MaxLength       =   6
         TabIndex        =   157
         Text            =   " "
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text124 
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
         Left            =   -71160
         MaxLength       =   6
         TabIndex        =   156
         Text            =   " "
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox Text123 
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
         Left            =   -72720
         MaxLength       =   6
         TabIndex        =   155
         Text            =   " "
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text122 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   154
         Text            =   " "
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Text121 
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
         MaxLength       =   6
         TabIndex        =   153
         Text            =   " "
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox Text120 
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
         Left            =   -66240
         MaxLength       =   20
         TabIndex        =   152
         Text            =   "4.5"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text119 
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
         MaxLength       =   20
         TabIndex        =   151
         Text            =   "6.55"
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox Text118 
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
         Left            =   -71160
         MaxLength       =   20
         TabIndex        =   150
         Text            =   "Descripciones Varias"
         Top             =   2040
         Width           =   3615
      End
      Begin VB.TextBox Text117 
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
         Left            =   -72720
         MaxLength       =   6
         TabIndex        =   149
         Text            =   " "
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox Text116 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   148
         Text            =   " "
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox Text115 
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
         MaxLength       =   6
         TabIndex        =   147
         Text            =   "O"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox Text114 
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
         Left            =   -66240
         MaxLength       =   20
         TabIndex        =   146
         Text            =   "3.25"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text113 
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
         MaxLength       =   20
         TabIndex        =   145
         Text            =   "10.25"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox Text112 
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
         Left            =   -71160
         MaxLength       =   50
         TabIndex        =   144
         Text            =   "Descripcion del Producto Terminado"
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text111 
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
         Left            =   -72720
         MaxLength       =   20
         TabIndex        =   143
         Text            =   "PT-12456-100"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox Text110 
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
         Left            =   -74160
         MaxLength       =   6
         TabIndex        =   142
         Text            =   " "
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text109 
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
         MaxLength       =   6
         TabIndex        =   141
         Text            =   "T"
         Top             =   1680
         Width           =   495
      End
      Begin VB.TextBox Text108 
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
         Left            =   -66240
         MaxLength       =   20
         TabIndex        =   134
         Text            =   "10.25"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text107 
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
         MaxLength       =   20
         TabIndex        =   133
         Text            =   "2.35"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text106 
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
         Left            =   -71160
         MaxLength       =   50
         TabIndex        =   132
         Text            =   "Descripcion de la Materia Prima"
         Top             =   1320
         Width           =   3615
      End
      Begin VB.TextBox Text105 
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
         Left            =   -72720
         MaxLength       =   20
         TabIndex        =   131
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text104 
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
         Left            =   -74160
         MaxLength       =   20
         TabIndex        =   130
         Text            =   "AA-026-100"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text103 
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
         MaxLength       =   6
         TabIndex        =   129
         Text            =   "M "
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text102 
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
         TabIndex        =   128
         Text            =   " "
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox Text101 
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
         TabIndex        =   127
         Text            =   " "
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox Text100 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   124
         Text            =   " "
         Top             =   4020
         Width           =   975
      End
      Begin VB.TextBox Text99 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   123
         Text            =   " "
         Top             =   4020
         Width           =   975
      End
      Begin VB.TextBox Text98 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   122
         Text            =   " "
         Top             =   4020
         Width           =   4455
      End
      Begin VB.TextBox Text97 
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
         TabIndex        =   121
         Text            =   " "
         Top             =   4020
         Width           =   855
      End
      Begin VB.TextBox Text96 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   120
         Text            =   " "
         Top             =   3660
         Width           =   975
      End
      Begin VB.TextBox Text95 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   119
         Text            =   " "
         Top             =   3660
         Width           =   975
      End
      Begin VB.TextBox Text94 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   118
         Text            =   " "
         Top             =   3660
         Width           =   4455
      End
      Begin VB.TextBox Text93 
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
         TabIndex        =   117
         Text            =   " "
         Top             =   3660
         Width           =   855
      End
      Begin VB.TextBox Text92 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   116
         Text            =   " "
         Top             =   3300
         Width           =   975
      End
      Begin VB.TextBox Text91 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   115
         Text            =   " "
         Top             =   3300
         Width           =   975
      End
      Begin VB.TextBox Text90 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   114
         Text            =   " "
         Top             =   3300
         Width           =   4455
      End
      Begin VB.TextBox Text89 
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
         TabIndex        =   113
         Text            =   " "
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox Text88 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   112
         Text            =   " "
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox Text87 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   111
         Text            =   " "
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox Text86 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   110
         Text            =   " "
         Top             =   2940
         Width           =   4455
      End
      Begin VB.TextBox Text85 
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
         TabIndex        =   109
         Text            =   " "
         Top             =   2940
         Width           =   855
      End
      Begin VB.TextBox Text84 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   108
         Text            =   " "
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox Text83 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   107
         Text            =   " "
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox Text82 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   106
         Text            =   " "
         Top             =   2580
         Width           =   4455
      End
      Begin VB.TextBox Text81 
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
         TabIndex        =   105
         Text            =   " "
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox Text80 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   104
         Text            =   " "
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox Text79 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   103
         Text            =   " "
         Top             =   2220
         Width           =   975
      End
      Begin VB.TextBox Text78 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   102
         Text            =   " "
         Top             =   2220
         Width           =   4455
      End
      Begin VB.TextBox Text77 
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
         TabIndex        =   101
         Text            =   " "
         Top             =   2220
         Width           =   855
      End
      Begin VB.TextBox Text76 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   100
         Text            =   " "
         Top             =   1860
         Width           =   975
      End
      Begin VB.TextBox Text75 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   99
         Text            =   " "
         Top             =   1860
         Width           =   975
      End
      Begin VB.TextBox Text74 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   98
         Text            =   " "
         Top             =   1860
         Width           =   4455
      End
      Begin VB.TextBox Text17 
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
         TabIndex        =   97
         Text            =   " "
         Top             =   1860
         Width           =   855
      End
      Begin VB.TextBox Text16 
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
         Left            =   -65040
         MaxLength       =   50
         TabIndex        =   90
         Text            =   " "
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox Text15 
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
         Left            =   -66120
         MaxLength       =   50
         TabIndex        =   89
         Text            =   " "
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox Text14 
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
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   88
         Text            =   " "
         Top             =   1500
         Width           =   855
      End
      Begin VB.TextBox Text13 
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
         Left            =   -68160
         MaxLength       =   50
         TabIndex        =   87
         Text            =   " "
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox Text12 
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
         Left            =   -69240
         MaxLength       =   50
         TabIndex        =   86
         Text            =   " "
         Top             =   1500
         Width           =   975
      End
      Begin VB.TextBox Text10 
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
         Left            =   -73800
         MaxLength       =   50
         TabIndex        =   85
         Text            =   " "
         Top             =   1500
         Width           =   4455
      End
      Begin VB.TextBox Text73 
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
         Left            =   -65400
         MaxLength       =   6
         TabIndex        =   84
         Text            =   " "
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox Text72 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   83
         Text            =   " "
         Top             =   3540
         Width           =   1095
      End
      Begin VB.TextBox Text71 
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
         Left            =   -71040
         MaxLength       =   6
         TabIndex        =   82
         Text            =   " "
         Top             =   3540
         Width           =   4335
      End
      Begin VB.TextBox Text70 
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
         Left            =   -72600
         MaxLength       =   6
         TabIndex        =   81
         Text            =   " "
         Top             =   3540
         Width           =   1455
      End
      Begin VB.TextBox Text69 
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
         MaxLength       =   6
         TabIndex        =   80
         Text            =   " "
         Top             =   3540
         Width           =   1335
      End
      Begin VB.TextBox Text68 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   79
         Text            =   " "
         Top             =   3540
         Width           =   495
      End
      Begin VB.TextBox Text67 
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
         Left            =   -65400
         MaxLength       =   6
         TabIndex        =   78
         Text            =   " "
         Top             =   3180
         Width           =   1095
      End
      Begin VB.TextBox Text66 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   77
         Text            =   " "
         Top             =   3180
         Width           =   1095
      End
      Begin VB.TextBox Text65 
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
         Left            =   -71040
         MaxLength       =   6
         TabIndex        =   76
         Text            =   " "
         Top             =   3180
         Width           =   4335
      End
      Begin VB.TextBox Text64 
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
         Left            =   -72600
         MaxLength       =   6
         TabIndex        =   75
         Text            =   " "
         Top             =   3180
         Width           =   1455
      End
      Begin VB.TextBox Text63 
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
         MaxLength       =   6
         TabIndex        =   74
         Text            =   " "
         Top             =   3180
         Width           =   1335
      End
      Begin VB.TextBox Text62 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   73
         Text            =   " "
         Top             =   3180
         Width           =   495
      End
      Begin VB.TextBox Text61 
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
         Left            =   -65400
         MaxLength       =   6
         TabIndex        =   72
         Text            =   " "
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox Text60 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   71
         Text            =   " "
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox Text59 
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
         Left            =   -71040
         MaxLength       =   6
         TabIndex        =   70
         Text            =   " "
         Top             =   2820
         Width           =   4335
      End
      Begin VB.TextBox Text58 
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
         Left            =   -72600
         MaxLength       =   6
         TabIndex        =   69
         Text            =   " "
         Top             =   2820
         Width           =   1455
      End
      Begin VB.TextBox Text57 
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
         MaxLength       =   6
         TabIndex        =   68
         Text            =   " "
         Top             =   2820
         Width           =   1335
      End
      Begin VB.TextBox Text56 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   67
         Text            =   " "
         Top             =   2820
         Width           =   495
      End
      Begin VB.TextBox Text55 
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
         Left            =   -65400
         MaxLength       =   6
         TabIndex        =   66
         Text            =   " "
         Top             =   2460
         Width           =   1095
      End
      Begin VB.TextBox Text54 
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
         Left            =   -66600
         MaxLength       =   6
         TabIndex        =   65
         Text            =   " "
         Top             =   2460
         Width           =   1095
      End
      Begin VB.TextBox Text53 
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
         Left            =   -71040
         MaxLength       =   6
         TabIndex        =   64
         Text            =   " "
         Top             =   2460
         Width           =   4335
      End
      Begin VB.TextBox Text52 
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
         Left            =   -72600
         MaxLength       =   6
         TabIndex        =   63
         Text            =   " "
         Top             =   2460
         Width           =   1455
      End
      Begin VB.TextBox Text51 
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
         MaxLength       =   6
         TabIndex        =   62
         Text            =   " "
         Top             =   2460
         Width           =   1335
      End
      Begin VB.TextBox Text50 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   61
         Text            =   " "
         Top             =   2460
         Width           =   495
      End
      Begin VB.TextBox Text49 
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
         Left            =   -65400
         MaxLength       =   20
         TabIndex        =   60
         Text            =   " "
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox Text48 
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
         Left            =   -66600
         MaxLength       =   20
         TabIndex        =   59
         Text            =   "6.55"
         Top             =   2100
         Width           =   1095
      End
      Begin VB.TextBox Text47 
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
         Left            =   -71040
         MaxLength       =   20
         TabIndex        =   58
         Text            =   "Descripciones Varias"
         Top             =   2100
         Width           =   4335
      End
      Begin VB.TextBox Text46 
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
         Left            =   -72600
         MaxLength       =   6
         TabIndex        =   57
         Text            =   " "
         Top             =   2100
         Width           =   1455
      End
      Begin VB.TextBox Text45 
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
         MaxLength       =   6
         TabIndex        =   56
         Text            =   " "
         Top             =   2100
         Width           =   1335
      End
      Begin VB.TextBox Text44 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   55
         Text            =   "O"
         Top             =   2100
         Width           =   495
      End
      Begin VB.TextBox Text43 
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
         Left            =   -65400
         MaxLength       =   20
         TabIndex        =   54
         Text            =   "852411"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text25 
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
         Left            =   -66600
         MaxLength       =   20
         TabIndex        =   53
         Text            =   "10.25"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.TextBox Text24 
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
         Left            =   -71040
         MaxLength       =   50
         TabIndex        =   52
         Text            =   "Descripcion del Producto Terminado"
         Top             =   1740
         Width           =   4335
      End
      Begin VB.TextBox Text23 
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
         Left            =   -72600
         MaxLength       =   20
         TabIndex        =   51
         Text            =   "PT-12456-100"
         Top             =   1740
         Width           =   1455
      End
      Begin VB.TextBox Text9 
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
         MaxLength       =   6
         TabIndex        =   50
         Text            =   " "
         Top             =   1740
         Width           =   1335
      End
      Begin VB.TextBox Text8 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   49
         Text            =   "T"
         Top             =   1740
         Width           =   495
      End
      Begin VB.TextBox Text7 
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
         Left            =   -65400
         MaxLength       =   20
         TabIndex        =   42
         Text            =   "101254"
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox Text6 
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
         Left            =   -66600
         MaxLength       =   20
         TabIndex        =   41
         Text            =   "2.35"
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox Text5 
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
         Left            =   -71040
         MaxLength       =   50
         TabIndex        =   40
         Text            =   "Descripcion de la Materia Prima"
         Top             =   1380
         Width           =   4335
      End
      Begin VB.TextBox Text4 
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
         Left            =   -72600
         MaxLength       =   20
         TabIndex        =   39
         Text            =   " "
         Top             =   1380
         Width           =   1455
      End
      Begin VB.TextBox Text3 
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
         MaxLength       =   20
         TabIndex        =   38
         Text            =   "AA-026-100"
         Top             =   1380
         Width           =   1335
      End
      Begin VB.TextBox Text2 
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
         Left            =   -74640
         MaxLength       =   6
         TabIndex        =   37
         Text            =   "M "
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox Text41 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   34
         Text            =   " "
         Top             =   4500
         Width           =   5895
      End
      Begin VB.TextBox Text40 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   33
         Text            =   " "
         Top             =   4140
         Width           =   5895
      End
      Begin VB.TextBox Text39 
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
         Left            =   -73080
         MaxLength       =   50
         TabIndex        =   31
         Text            =   " "
         Top             =   3780
         Width           =   5895
      End
      Begin VB.TextBox Text38 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   30
         Text            =   " "
         Top             =   3300
         Width           =   4575
      End
      Begin VB.TextBox Text37 
         Height          =   285
         Left            =   -74760
         TabIndex        =   29
         Top             =   3300
         Width           =   735
      End
      Begin VB.TextBox Text36 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   28
         Text            =   " "
         Top             =   3300
         Width           =   4335
      End
      Begin VB.TextBox Text35 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   27
         Text            =   " "
         Top             =   2940
         Width           =   4575
      End
      Begin VB.TextBox Text34 
         Height          =   285
         Left            =   -74760
         TabIndex        =   26
         Top             =   2940
         Width           =   735
      End
      Begin VB.TextBox Text33 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   25
         Text            =   " "
         Top             =   2940
         Width           =   4335
      End
      Begin VB.TextBox Text32 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   24
         Text            =   " "
         Top             =   2580
         Width           =   4575
      End
      Begin VB.TextBox Text31 
         Height          =   285
         Left            =   -74760
         TabIndex        =   23
         Top             =   2580
         Width           =   735
      End
      Begin VB.TextBox Text30 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   22
         Text            =   " "
         Top             =   2580
         Width           =   4335
      End
      Begin VB.TextBox Text29 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   21
         Text            =   " "
         Top             =   2220
         Width           =   4575
      End
      Begin VB.TextBox Text28 
         Height          =   285
         Left            =   -74760
         TabIndex        =   20
         Top             =   2220
         Width           =   735
      End
      Begin VB.TextBox Text27 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   19
         Text            =   " "
         Top             =   2220
         Width           =   4335
      End
      Begin VB.TextBox Text22 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   18
         Text            =   " "
         Top             =   1860
         Width           =   4575
      End
      Begin VB.TextBox Text21 
         Height          =   285
         Left            =   -74760
         TabIndex        =   17
         Top             =   1860
         Width           =   735
      End
      Begin VB.TextBox Text20 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   16
         Text            =   " "
         Top             =   1860
         Width           =   4335
      End
      Begin VB.TextBox Text19 
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
         Left            =   -69480
         MaxLength       =   50
         TabIndex        =   12
         Text            =   " "
         Top             =   1500
         Width           =   4575
      End
      Begin VB.TextBox Text26 
         Height          =   285
         Left            =   -74760
         TabIndex        =   10
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Text18 
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
         Left            =   -73920
         MaxLength       =   50
         TabIndex        =   9
         Text            =   " "
         Top             =   1500
         Width           =   4335
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   8
         Text            =   " "
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label35 
         Caption         =   "Observaciones"
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
         Left            =   360
         TabIndex        =   224
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label Label34 
         Caption         =   "Controles Adicionales"
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
         Left            =   -74640
         TabIndex        =   212
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label33 
         Caption         =   "Cantidad"
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
         Left            =   -66840
         TabIndex        =   188
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label32 
         Caption         =   "Cantidad"
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
         Left            =   -66840
         TabIndex        =   187
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label31 
         Caption         =   "Importe"
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
         Left            =   -65040
         TabIndex        =   178
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "Cantidad"
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
         Left            =   -67440
         TabIndex        =   140
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label29 
         Caption         =   "Costo"
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
         Left            =   -66240
         TabIndex        =   139
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label28 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71160
         TabIndex        =   138
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Terminado"
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
         Left            =   -72720
         TabIndex        =   137
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Articulo"
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
         Left            =   -74160
         TabIndex        =   136
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label14 
         Caption         =   "Tipo"
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
         TabIndex        =   135
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label27 
         Caption         =   "Encargado del Proyecto"
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
         Left            =   360
         TabIndex        =   126
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label Label26 
         Caption         =   "Preparo"
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
         Left            =   360
         TabIndex        =   125
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "Instrucciones Seguridad"
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
         Height          =   495
         Left            =   -65040
         TabIndex        =   96
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label Label24 
         Caption         =   "Control"
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
         Left            =   -66120
         TabIndex        =   95
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label23 
         Caption         =   "Tiempo"
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
         Left            =   -67080
         TabIndex        =   94
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Temperatura"
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
         Height          =   495
         Left            =   -68160
         TabIndex        =   93
         Top             =   900
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "Equipo"
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
         Left            =   -69240
         TabIndex        =   92
         Top             =   900
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Instrucciones de Trabajo"
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
         Left            =   -73800
         TabIndex        =   91
         Top             =   900
         Width           =   3015
      End
      Begin VB.Label Label18 
         Caption         =   "Cantidad"
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
         Left            =   -66600
         TabIndex        =   48
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Lote"
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
         Left            =   -65400
         TabIndex        =   47
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label8 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -71040
         TabIndex        =   46
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Terminado"
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
         Left            =   -72600
         TabIndex        =   45
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Articulo"
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
         Left            =   -74040
         TabIndex        =   44
         Top             =   1020
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
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
         Left            =   -74640
         TabIndex        =   43
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label22 
         Caption         =   "Observaciones"
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
         Left            =   -74760
         TabIndex        =   32
         Top             =   3780
         Width           =   1815
      End
      Begin VB.Label Label21 
         Caption         =   "Resultado Obtenido"
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
         Left            =   -69480
         TabIndex        =   15
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label Label20 
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
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   -73920
         TabIndex        =   14
         Top             =   1140
         Width           =   2175
      End
      Begin VB.Label Label19 
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
         Height          =   285
         Left            =   -74760
         TabIndex        =   13
         Top             =   1140
         Width           =   735
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "ESPECIFICACIONES"
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
         Left            =   -74520
         TabIndex        =   11
         Top             =   900
         Width           =   6975
      End
      Begin VB.Label Label9 
         Caption         =   "Etapa"
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
         Left            =   -74760
         TabIndex        =   7
         Top             =   900
         Width           =   855
      End
   End
   Begin VB.TextBox Cliente 
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
      MaxLength       =   6
      TabIndex        =   4
      Text            =   " "
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox Text1 
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
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin MSMask.MaskEdBox Fecha 
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   240
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
   Begin VB.Label Label3 
      Caption         =   "Cantidad "
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
      Left            =   3840
      TabIndex        =   36
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Height          =   285
      Left            =   360
      TabIndex        =   5
      Top             =   600
      Width           =   1575
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
      Left            =   3720
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Orden de Trabajo"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "DemoDesarrolloII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
