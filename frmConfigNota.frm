VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigNota 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuração de nota fiscal"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   183
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Data dbConfigNota 
      Caption         =   "dbConfigNota"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select *from ConfigNota"
      Top             =   5640
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Natureza da Operação"
      TabPicture(0)   =   "frmConfigNota.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label76"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label77"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Text1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Text2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Text3"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Text4"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text6"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Text7"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text10"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text11"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text12"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text13"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text14"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text15"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text16"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text102"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text103"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text104"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text105"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "Destinatário"
      TabPicture(1)   =   "frmConfigNota.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(8)=   "Label20"
      Tab(1).Control(9)=   "Label21"
      Tab(1).Control(10)=   "Label22"
      Tab(1).Control(11)=   "Label23"
      Tab(1).Control(12)=   "Text17"
      Tab(1).Control(13)=   "Text18"
      Tab(1).Control(14)=   "Text19"
      Tab(1).Control(15)=   "Text20"
      Tab(1).Control(16)=   "Text21"
      Tab(1).Control(17)=   "Text22"
      Tab(1).Control(18)=   "Text23"
      Tab(1).Control(19)=   "Text24"
      Tab(1).Control(20)=   "Text25"
      Tab(1).Control(21)=   "Text26"
      Tab(1).Control(22)=   "Text27"
      Tab(1).Control(23)=   "Text28"
      Tab(1).Control(24)=   "Text29"
      Tab(1).Control(25)=   "Text30"
      Tab(1).Control(26)=   "Text31"
      Tab(1).Control(27)=   "Text32"
      Tab(1).Control(28)=   "Text33"
      Tab(1).Control(29)=   "Text34"
      Tab(1).ControlCount=   30
      TabCaption(2)   =   "Corpo"
      TabPicture(2)   =   "frmConfigNota.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text112"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Text111"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Text110"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Text109"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Text108"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Text107"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Text106"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Text67"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Text66"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Text65"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "Text64"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "Text63"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "Text62"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "Text61"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "Text60"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "Text59"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "Text58"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "Text57"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "Text56"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "Text55"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "Text54"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "Text53"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "Text52"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "Text51"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "Text50"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "Text49"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "Text48"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "Text47"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "Text46"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "Text45"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "Text44"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "Text43"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "Text42"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "Text41"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "Text40"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "Text39"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "Text38"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "Text37"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "Text36"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "Text35"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "Label81"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "Label80"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "Label79"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "Label78"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "Label49"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "Label48"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "Label47"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "Label46"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "Label45"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "Label44"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "Label43"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "Label42"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "Label41"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "Label40"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Label39"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Label38"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "Label37"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "Label36"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "Label35"
      Tab(2).Control(58).Enabled=   0   'False
      Tab(2).Control(59)=   "Label34"
      Tab(2).Control(59).Enabled=   0   'False
      Tab(2).Control(60)=   "Label33"
      Tab(2).Control(60).Enabled=   0   'False
      Tab(2).Control(61)=   "Label32"
      Tab(2).Control(61).Enabled=   0   'False
      Tab(2).Control(62)=   "Label31"
      Tab(2).Control(62).Enabled=   0   'False
      Tab(2).Control(63)=   "Label30"
      Tab(2).Control(63).Enabled=   0   'False
      Tab(2).Control(64)=   "Label29"
      Tab(2).Control(64).Enabled=   0   'False
      Tab(2).Control(65)=   "Label28"
      Tab(2).Control(65).Enabled=   0   'False
      Tab(2).Control(66)=   "Label27"
      Tab(2).Control(66).Enabled=   0   'False
      Tab(2).Control(67)=   "Label26"
      Tab(2).Control(67).Enabled=   0   'False
      Tab(2).Control(68)=   "Label25"
      Tab(2).Control(68).Enabled=   0   'False
      Tab(2).Control(69)=   "Label24"
      Tab(2).Control(69).Enabled=   0   'False
      Tab(2).ControlCount=   70
      TabCaption(3)   =   "Transportador"
      TabPicture(3)   =   "frmConfigNota.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label50"
      Tab(3).Control(1)=   "Label51"
      Tab(3).Control(2)=   "Label52"
      Tab(3).Control(3)=   "Label53"
      Tab(3).Control(4)=   "Label54"
      Tab(3).Control(5)=   "Label55"
      Tab(3).Control(6)=   "Label56"
      Tab(3).Control(7)=   "Label57"
      Tab(3).Control(8)=   "Label58"
      Tab(3).Control(9)=   "Label59"
      Tab(3).Control(10)=   "Label60"
      Tab(3).Control(11)=   "Label61"
      Tab(3).Control(12)=   "Label62"
      Tab(3).Control(13)=   "Label63"
      Tab(3).Control(14)=   "Label64"
      Tab(3).Control(15)=   "Label65"
      Tab(3).Control(16)=   "Label66"
      Tab(3).Control(17)=   "Label67"
      Tab(3).Control(18)=   "Label68"
      Tab(3).Control(19)=   "Label69"
      Tab(3).Control(20)=   "Label70"
      Tab(3).Control(21)=   "Label71"
      Tab(3).Control(22)=   "Label72"
      Tab(3).Control(23)=   "Label73"
      Tab(3).Control(24)=   "Label74"
      Tab(3).Control(25)=   "Label75"
      Tab(3).Control(26)=   "Text68"
      Tab(3).Control(27)=   "Text69"
      Tab(3).Control(28)=   "Text70"
      Tab(3).Control(29)=   "Text71"
      Tab(3).Control(30)=   "Text72"
      Tab(3).Control(31)=   "Text73"
      Tab(3).Control(32)=   "Text74"
      Tab(3).Control(33)=   "Text75"
      Tab(3).Control(34)=   "Text76"
      Tab(3).Control(35)=   "Text77"
      Tab(3).Control(36)=   "Text78"
      Tab(3).Control(37)=   "Text79"
      Tab(3).Control(38)=   "Text80"
      Tab(3).Control(39)=   "Text81"
      Tab(3).Control(40)=   "Text82"
      Tab(3).Control(41)=   "Text83"
      Tab(3).Control(42)=   "Text84"
      Tab(3).Control(43)=   "Text85"
      Tab(3).Control(44)=   "Text86"
      Tab(3).Control(45)=   "Text87"
      Tab(3).Control(46)=   "Text88"
      Tab(3).Control(47)=   "Text89"
      Tab(3).Control(48)=   "Text90"
      Tab(3).Control(49)=   "Text91"
      Tab(3).Control(50)=   "Text92"
      Tab(3).Control(51)=   "Text93"
      Tab(3).Control(52)=   "Text94"
      Tab(3).Control(53)=   "Text95"
      Tab(3).Control(54)=   "Text96"
      Tab(3).Control(55)=   "Text97"
      Tab(3).Control(56)=   "Text98"
      Tab(3).Control(57)=   "Text99"
      Tab(3).Control(58)=   "Text100"
      Tab(3).Control(59)=   "Text101"
      Tab(3).ControlCount=   60
      Begin VB.TextBox Text112 
         Alignment       =   1  'Right Justify
         DataField       =   "LinhasCorpo"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   193
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Text111 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   189
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox Text110 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   188
         Top             =   4320
         Width           =   855
      End
      Begin VB.TextBox Text109 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoISSX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   187
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Text108 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoISSY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   186
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Text107 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoTotalX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   185
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text106 
         Alignment       =   1  'Right Justify
         DataField       =   "PrestacaoServicoTotalY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   184
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Text105 
         Alignment       =   1  'Right Justify
         DataField       =   "NrNotaCanhotoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   8
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text104 
         Alignment       =   1  'Right Justify
         DataField       =   "NrNotaCanhotoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text103 
         Alignment       =   1  'Right Justify
         DataField       =   "NrNotaTopoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text102 
         Alignment       =   1  'Right Justify
         DataField       =   "NrNotaTopoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text101 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosAdicionais2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -69840
         TabIndex        =   181
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox Text100 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosAdicionais2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -70800
         TabIndex        =   180
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox Text99 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosAdicionais1Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -69840
         TabIndex        =   178
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text98 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosAdicionais1X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -70800
         TabIndex        =   177
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text97 
         Alignment       =   1  'Right Justify
         DataField       =   "PesoLiquidoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   172
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text96 
         Alignment       =   1  'Right Justify
         DataField       =   "PesoLiquidoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   171
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text95 
         Alignment       =   1  'Right Justify
         DataField       =   "PesoBrutoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   169
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text94 
         Alignment       =   1  'Right Justify
         DataField       =   "PesoBrutoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   168
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text93 
         Alignment       =   1  'Right Justify
         DataField       =   "NumeroY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   166
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text92 
         Alignment       =   1  'Right Justify
         DataField       =   "NumeroX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   165
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text91 
         Alignment       =   1  'Right Justify
         DataField       =   "EspecieY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   160
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text90 
         Alignment       =   1  'Right Justify
         DataField       =   "EspecieX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   159
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text89 
         Alignment       =   1  'Right Justify
         DataField       =   "MarcaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   163
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text88 
         Alignment       =   1  'Right Justify
         DataField       =   "MarcaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   162
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text87 
         Alignment       =   1  'Right Justify
         DataField       =   "QTD2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   157
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text86 
         Alignment       =   1  'Right Justify
         DataField       =   "QTD2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   156
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text85 
         Alignment       =   1  'Right Justify
         DataField       =   "IE2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67680
         TabIndex        =   154
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text84 
         Alignment       =   1  'Right Justify
         DataField       =   "IE2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68640
         TabIndex        =   153
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text83 
         Alignment       =   1  'Right Justify
         DataField       =   "Municipio2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   145
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text82 
         Alignment       =   1  'Right Justify
         DataField       =   "Municipio2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   144
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text81 
         Alignment       =   1  'Right Justify
         DataField       =   "Endereco2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   142
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text80 
         Alignment       =   1  'Right Justify
         DataField       =   "Endereco2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   141
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text79 
         Alignment       =   1  'Right Justify
         DataField       =   "CNPJ2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   139
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text78 
         Alignment       =   1  'Right Justify
         DataField       =   "CNPJ2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   138
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text77 
         Alignment       =   1  'Right Justify
         DataField       =   "UF3Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   148
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text76 
         Alignment       =   1  'Right Justify
         DataField       =   "UF3X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   147
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text75 
         Alignment       =   1  'Right Justify
         DataField       =   "UF2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   136
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text74 
         Alignment       =   1  'Right Justify
         DataField       =   "UF2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   135
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text73 
         Alignment       =   1  'Right Justify
         DataField       =   "PlacaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   133
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text72 
         Alignment       =   1  'Right Justify
         DataField       =   "PlacaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   132
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text71 
         Alignment       =   1  'Right Justify
         DataField       =   "FretePorContaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   130
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text70 
         Alignment       =   1  'Right Justify
         DataField       =   "FretePorContaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   129
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text69 
         Alignment       =   1  'Right Justify
         DataField       =   "Nome2Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71760
         TabIndex        =   127
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text68 
         Alignment       =   1  'Right Justify
         DataField       =   "Nome2X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72720
         TabIndex        =   126
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text67 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalNotaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   121
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text66 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalNotaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   120
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text65 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalIPIY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   118
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text64 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalIPIX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   117
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text63 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorDoSeguroY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   112
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text62 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorDoSeguroX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   111
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text61 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorDoFreteY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   109
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text60 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorDoFreteX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   108
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text59 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalProdutosY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   106
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text58 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorTotalProdutosX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   105
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text57 
         Alignment       =   1  'Right Justify
         DataField       =   "OutrasDespY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   115
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text56 
         Alignment       =   1  'Right Justify
         DataField       =   "OutrasDespX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   114
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text55 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMSSubY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   103
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text54 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMSSubX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   102
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text53 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMSSubY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   100
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text52 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMSSubX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   99
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text51 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMSY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   97
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text50 
         Alignment       =   1  'Right Justify
         DataField       =   "ValorICMSX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   96
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text49 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMSY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -67560
         TabIndex        =   94
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text48 
         Alignment       =   1  'Right Justify
         DataField       =   "BaseICMSX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -68520
         TabIndex        =   93
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text47 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaLimite"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   88
         Top             =   4920
         Width           =   855
      End
      Begin VB.TextBox Text46 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaValorIPI"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   86
         Top             =   4560
         Width           =   855
      End
      Begin VB.TextBox Text45 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaAliquotaIPI"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   84
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text44 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaAliquotaICMS"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   82
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox Text43 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaVTotal"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   80
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox Text42 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaVUnitario"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   78
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox Text41 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaQuantidade"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   76
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text40 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaUnidade"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   74
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox Text39 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaSubstTrib"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   72
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text38 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaClasFiscal"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   70
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox Text37 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaDescri"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   68
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text36 
         Alignment       =   1  'Right Justify
         DataField       =   "ColunaCodigo"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text35 
         Alignment       =   1  'Right Justify
         DataField       =   "InicioCorpoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71520
         TabIndex        =   64
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text34 
         Alignment       =   1  'Right Justify
         DataField       =   "IEY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   62
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text33 
         Alignment       =   1  'Right Justify
         DataField       =   "IEX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   61
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text32 
         Alignment       =   1  'Right Justify
         DataField       =   "FoneY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   56
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text31 
         Alignment       =   1  'Right Justify
         DataField       =   "FoneX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   55
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text30 
         Alignment       =   1  'Right Justify
         DataField       =   "MunicipioY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   53
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text29 
         Alignment       =   1  'Right Justify
         DataField       =   "MunicipioX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   52
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text28 
         Alignment       =   1  'Right Justify
         DataField       =   "CEPY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   50
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text27 
         Alignment       =   1  'Right Justify
         DataField       =   "CEPX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   49
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text26 
         Alignment       =   1  'Right Justify
         DataField       =   "UF1Y"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   59
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
         DataField       =   "UF1X"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   58
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text24 
         Alignment       =   1  'Right Justify
         DataField       =   "BairroY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   47
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         DataField       =   "BairroX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   46
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
         DataField       =   "EnderecoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   44
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text21 
         Alignment       =   1  'Right Justify
         DataField       =   "EnderecoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   43
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         DataField       =   "CNPJY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   41
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         DataField       =   "CNPJX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   40
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         DataField       =   "NomeY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -71160
         TabIndex        =   38
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         DataField       =   "NomeX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   -72120
         TabIndex        =   37
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         DataField       =   "HoraSaidaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   29
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         DataField       =   "HoraSaidaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   28
         Top             =   3600
         Width           =   855
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         DataField       =   "DataSaidaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   26
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         DataField       =   "DataSaidaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   25
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         DataField       =   "DataEmissaoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   23
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         DataField       =   "DataEmissaoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosFaturaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   32
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         DataField       =   "DadosFaturaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   31
         Top             =   3960
         Width           =   855
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         DataField       =   "CFOPY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         DataField       =   "CFOPX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         DataField       =   "NaturezaOperacaoY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   17
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         DataField       =   "NaturezaOperacaoX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   16
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         DataField       =   "EntradaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         DataField       =   "EntradaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   13
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         DataField       =   "SaidaY"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   3840
         TabIndex        =   11
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         DataField       =   "SaidaX"
         DataSource      =   "dbConfigNota"
         Height          =   285
         Left            =   2880
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label81 
         Alignment       =   1  'Right Justify
         Caption         =   "Limite de linhas para impressão na nota:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   194
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label Label80 
         Alignment       =   1  'Right Justify
         Caption         =   "Prestação de Serviços:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   192
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label79 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor do ISS:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   191
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label Label78 
         Alignment       =   1  'Right Justify
         Caption         =   "Total dos Serviços:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   190
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label77 
         Alignment       =   1  'Right Justify
         Caption         =   "Número da nota no Canhoto:"
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label76 
         Alignment       =   1  'Right Justify
         Caption         =   "Número da nota no topo:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label75 
         Caption         =   "Y"
         Height          =   255
         Left            =   -69840
         TabIndex        =   175
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label74 
         Caption         =   "X"
         Height          =   255
         Left            =   -70800
         TabIndex        =   174
         Top             =   4320
         Width           =   255
      End
      Begin VB.Label Label73 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   -72840
         TabIndex        =   173
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label Label72 
         Alignment       =   1  'Right Justify
         Caption         =   "Limite para os Dados Adicionais:"
         Height          =   255
         Left            =   -73440
         TabIndex        =   179
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         Caption         =   "Início dos Dados Adicionais:"
         Height          =   255
         Left            =   -73440
         TabIndex        =   176
         Top             =   4560
         Width           =   2535
      End
      Begin VB.Label Label70 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso Líquido:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   170
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label69 
         Caption         =   "Y"
         Height          =   255
         Left            =   -67680
         TabIndex        =   151
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label68 
         Caption         =   "X"
         Height          =   255
         Left            =   -68640
         TabIndex        =   150
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   149
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label66 
         Alignment       =   1  'Right Justify
         Caption         =   "Peso Bruto:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   167
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         Caption         =   "Número:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   164
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         Caption         =   "Espécie:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   158
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         Caption         =   "Marca:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   161
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         Caption         =   "Quantidade:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   155
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         Caption         =   "Inscrição Estadual:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   152
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         Caption         =   "Município:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   143
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   140
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   137
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   146
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   134
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         Caption         =   "Placa:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   131
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         Caption         =   "Frete por conta:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   128
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label53 
         Caption         =   "Y"
         Height          =   255
         Left            =   -71760
         TabIndex        =   124
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label52 
         Caption         =   "X"
         Height          =   255
         Left            =   -72720
         TabIndex        =   123
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   122
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome/Razão Social:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   125
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor total da Nota:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   119
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor total do IPI:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   116
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor do Seguro:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   110
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor do Frete:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   107
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor total dos produtos:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   104
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         Caption         =   "Outras Desp. Acessórias:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   113
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor ICMS Subst.:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   101
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "B. de Calc ICMS Subst.:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   98
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor ICMS:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   95
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label40 
         Caption         =   "Y"
         Height          =   255
         Left            =   -67560
         TabIndex        =   91
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label39 
         Caption         =   "X"
         Height          =   255
         Left            =   -68520
         TabIndex        =   90
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   89
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label37 
         Alignment       =   1  'Right Justify
         Caption         =   "Base de Cálculo ICMS:"
         Height          =   255
         Left            =   -70560
         TabIndex        =   92
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Limite X para impressão na nota:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   87
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Valor IPI:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   85
         Top             =   4560
         Width           =   3135
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Alíquota IPI:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   83
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Alíquota ICMS:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   81
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Valor Total:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         Top             =   3480
         Width           =   3135
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Valor Unitário:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   77
         Top             =   3120
         Width           =   3135
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Quantidade:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   75
         Top             =   2760
         Width           =   3135
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Unidade:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   73
         Top             =   2400
         Width           =   3135
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Subst. Tributária:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   71
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Clas. Fiscal:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   1680
         Width           =   3135
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Descrição:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Caption         =   "Início X da Coluna Código:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   65
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenada Y para início do Corpo da nota:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   63
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Inscrição Estadual:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   60
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Fone/Fax:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   54
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Município:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   51
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   48
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   57
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro/Distrito:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   45
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   42
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "C.N.P.J./C.P.F.:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   39
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label15 
         Caption         =   "Y"
         Height          =   255
         Left            =   -71160
         TabIndex        =   35
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "X"
         Height          =   255
         Left            =   -72120
         TabIndex        =   34
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   33
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome/Razão Social:"
         Height          =   255
         Left            =   -74640
         TabIndex        =   36
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Hora da Saída:"
         Height          =   255
         Left            =   360
         TabIndex        =   27
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Data da Saída/Entrada:"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Data da Emissão:"
         Height          =   255
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Dados da Fatura:"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "CFOP:"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Natureza da Operação:"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Entrada:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Y"
         Height          =   255
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "X"
         Height          =   255
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Coordenadas em cm:"
         Height          =   255
         Left            =   360
         TabIndex        =   182
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Saída:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1440
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmConfigNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
Dim db As Database, Ws As Workspace

With dbConfigNota
  .Connect = Conectar
  .DatabaseName = Caminho
  'On Error GoTo 0
  'On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    On Error GoTo 0
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho, , , Conectar)
    db.Execute "create table ConfigNota (NrNotaTopoX double)"
    'Topo da Nota
    db.Execute "alter table Confignota add column NrNotaTopoY double"
    db.Execute "alter table Confignota add column NrNotaCanhotoX double"
    db.Execute "alter table Confignota add column NrNotaCanhotoY double"
    db.Execute "alter table Confignota add column SaidaX double"
    db.Execute "alter table Confignota add column SaidaY double"
    db.Execute "alter table Confignota add column EntradaX double"
    db.Execute "alter table Confignota add column EntradaY double"
    db.Execute "alter table Confignota add column NaturezaOperacaoX double"
    db.Execute "alter table Confignota add column NaturezaOperacaoY double"
    db.Execute "alter table Confignota add column CFOPX double"
    db.Execute "alter table Confignota add column CFOPY double"
    db.Execute "alter table Confignota add column DataEmissaoX double"
    db.Execute "alter table Confignota add column DataEmissaoY double"
    db.Execute "alter table Confignota add column DataSaidaX double"
    db.Execute "alter table Confignota add column DataSaidaY double"
    db.Execute "alter table Confignota add column HoraSaidaX double"
    db.Execute "alter table Confignota add column HoraSaidaY double"
    db.Execute "alter table Confignota add column DadosFaturaX double"
    db.Execute "alter table Confignota add column DadosFaturaY double"
    'Destinatário
    db.Execute "alter table Confignota add column NomeX double"
    db.Execute "alter table Confignota add column NomeY double"
    db.Execute "alter table Confignota add column CNPJX double"
    db.Execute "alter table Confignota add column CNPJY double"
    db.Execute "alter table Confignota add column EnderecoX double"
    db.Execute "alter table Confignota add column EnderecoY double"
    db.Execute "alter table Confignota add column BairroX double"
    db.Execute "alter table Confignota add column BairroY double"
    db.Execute "alter table Confignota add column CEPX double"
    db.Execute "alter table Confignota add column CEPY double"
    db.Execute "alter table Confignota add column MunicipioX double"
    db.Execute "alter table Confignota add column MunicipioY double"
    db.Execute "alter table Confignota add column FoneX double"
    db.Execute "alter table Confignota add column FoneY double"
    db.Execute "alter table Confignota add column UF1X double"
    db.Execute "alter table Confignota add column UF1Y double"
    db.Execute "alter table Confignota add column IEX double"
    db.Execute "alter table Confignota add column IEY double"
    'Corpo
    db.Execute "alter table Confignota add column InicioCorpoY double"
    db.Execute "alter table Confignota add column ColunaCodigo double"
    db.Execute "alter table Confignota add column ColunaDescri double"
    db.Execute "alter table Confignota add column ColunaClasFiscal double"
    db.Execute "alter table Confignota add column ColunaSubstTrib double"
    db.Execute "alter table Confignota add column ColunaUnidade double"
    db.Execute "alter table Confignota add column ColunaQuantidade double"
    db.Execute "alter table Confignota add column ColunaVUnitario double"
    db.Execute "alter table Confignota add column ColunaVTotal double"
    db.Execute "alter table Confignota add column ColunaAliquotaICMS double"
    db.Execute "alter table Confignota add column ColunaAliquotaIPI double"
    db.Execute "alter table Confignota add column ColunaValorIPI double"
    db.Execute "alter table Confignota add column ColunaLimite double"
    db.Execute "alter table Confignota add column BaseICMSX double"
    db.Execute "alter table Confignota add column BaseICMSY double"
    db.Execute "alter table Confignota add column ValorICMSX double"
    db.Execute "alter table Confignota add column ValorICMSY double"
    db.Execute "alter table Confignota add column BaseICMSSubX double"
    db.Execute "alter table Confignota add column BaseICMSSubY double"
    db.Execute "alter table Confignota add column ValorICMSSubX double"
    db.Execute "alter table Confignota add column ValorICMSSubY double"
    db.Execute "alter table Confignota add column ValorTotalProdutosX double"
    db.Execute "alter table Confignota add column ValorTotalProdutosY double"
    db.Execute "alter table Confignota add column ValorDoFreteX double"
    db.Execute "alter table Confignota add column ValorDoFreteY double"
    db.Execute "alter table Confignota add column ValorDoSeguroX double"
    db.Execute "alter table Confignota add column ValorDoSeguroY double"
    db.Execute "alter table Confignota add column OutrasDespX double"
    db.Execute "alter table Confignota add column OutrasDespY double"
    db.Execute "alter table Confignota add column ValorTotalIPIX double"
    db.Execute "alter table Confignota add column ValorTotalIPIY double"
    db.Execute "alter table Confignota add column ValorTotalNotaX double"
    db.Execute "alter table Confignota add column ValorTotalNotaY double"
    'Transportador
    db.Execute "alter table Confignota add column Nome2X double"
    db.Execute "alter table Confignota add column Nome2Y double"
    db.Execute "alter table Confignota add column FretePorContaX double"
    db.Execute "alter table Confignota add column FretePorContaY double"
    db.Execute "alter table Confignota add column PlacaX double"
    db.Execute "alter table Confignota add column PlacaY double"
    db.Execute "alter table Confignota add column UF2X double"
    db.Execute "alter table Confignota add column UF2Y double"
    db.Execute "alter table Confignota add column CNPJ2X double"
    db.Execute "alter table Confignota add column CNPJ2Y double"
    db.Execute "alter table Confignota add column Endereco2X double"
    db.Execute "alter table Confignota add column Endereco2Y double"
    db.Execute "alter table Confignota add column Municipio2X double"
    db.Execute "alter table Confignota add column Municipio2Y double"
    db.Execute "alter table Confignota add column UF3X double"
    db.Execute "alter table Confignota add column UF3Y double"
    db.Execute "alter table Confignota add column IE2X double"
    db.Execute "alter table Confignota add column IE2Y double"
    db.Execute "alter table Confignota add column QTD2X double"
    db.Execute "alter table Confignota add column QTD2Y double"
    db.Execute "alter table Confignota add column EspecieX double"
    db.Execute "alter table Confignota add column EspecieY double"
    db.Execute "alter table Confignota add column MarcaX double"
    db.Execute "alter table Confignota add column MarcaY double"
    db.Execute "alter table Confignota add column NumeroX double"
    db.Execute "alter table Confignota add column NumeroY double"
    db.Execute "alter table Confignota add column PesoBrutoX double"
    db.Execute "alter table Confignota add column PesoBrutoY double"
    db.Execute "alter table Confignota add column PesoLiquidoX double"
    db.Execute "alter table Confignota add column PesoLiquidoY double"
    db.Execute "alter table Confignota add column DadosAdicionais1X double"
    db.Execute "alter table Confignota add column DadosAdicionais1Y double"
    db.Execute "alter table Confignota add column DadosAdicionais2X double"
    db.Execute "alter table Confignota add column DadosAdicionais2Y double"
  End If
End With
With dbConfigNota
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    For i = 0 To .Recordset.Fields.Count - 1
      .Recordset(i) = 0
    Next i
    .Recordset.Update
  End If
End With
End Sub

