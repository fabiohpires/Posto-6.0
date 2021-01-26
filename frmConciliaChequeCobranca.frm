VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaChequeCobranca 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobrança de Cheques Devolvidos"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmConciliaChequeCobranca.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   720
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from bloqueiafechamento"
      Caption         =   "dbBloqueiaFechamento"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data dbCarros 
      Caption         =   "dbCarros"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from estatus"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbStatus 
      Caption         =   "dbStatus"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from estatus"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbClientesContas 
      Caption         =   "dbClientesContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where cobrando=-1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "chequesclientes"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data qTotalPendente 
      Caption         =   "qTotalPendente"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where codigosoma='1' and compensado=0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data QSomaCheques 
      Caption         =   "QSomaCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where cobrando=-1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmConciliaChequeCobranca.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbDepositar 
      Caption         =   "dbDepositar"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0 and devolvido=-1 and protesto=0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ConciliaNova"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdExibe 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6360
      TabIndex        =   17
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   8760
      Picture         =   "frmConciliaChequeCobranca.frx":0459
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "Imprimir"
      Top             =   2040
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6360
      Top             =   840
   End
   Begin VB.CommandButton cmdSomar 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdSubtrair 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6360
      TabIndex        =   29
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   5280
      TabIndex        =   27
      Top             =   5400
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   3840
      TabIndex        =   25
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   37683
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5760
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   3
      Mask            =   "999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   4
      Mask            =   "9999"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   2280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   8
      Mask            =   "999999-9"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   2280
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   6
      Mask            =   "999999"
      PromptChar      =   " "
   End
   Begin MSComCtl2.DTPicker txtDataini 
      Height          =   285
      Left            =   3240
      TabIndex        =   14
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   37683
   End
   Begin MSComCtl2.DTPicker txtDatafim 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72286209
      CurrentDate     =   37683
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaChequeCobranca.frx":0EDB
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmConciliaChequeCobranca.frx":0EF3
      TabIndex        =   35
      Top             =   120
      Width           =   9255
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmConciliaChequeCobranca.frx":27FE
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmConciliaChequeCobranca.frx":2818
      TabIndex        =   36
      Top             =   3120
      Width           =   9255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   7320
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label70 
      Caption         =   "Leitura Automática"
      Height          =   255
      Left            =   7680
      TabIndex        =   34
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   4680
      TabIndex        =   15
      Top             =   2760
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período para exibir cheques em cobrança:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Valor Recebido:"
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   5280
      Width           =   1140
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   465
   End
   Begin VB.Label lblTotalPendente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   33
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Total Depositado:"
      Height          =   255
      Left            =   6000
      TabIndex        =   32
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   7800
      TabIndex        =   31
      Top             =   5520
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total em Cobrança:"
      Height          =   195
      Left            =   7800
      TabIndex        =   30
      Top             =   5280
      Width           =   1395
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   28
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   4920
      TabIndex        =   26
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
      Height          =   195
      Left            =   3840
      TabIndex        =   22
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      Height          =   195
      Left            =   3000
      TabIndex        =   8
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      Height          =   195
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      Caption         =   "Comp:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data Pgto.:"
      Height          =   195
      Left            =   3840
      TabIndex        =   23
      Top             =   5280
      Width           =   810
   End
End
Attribute VB_Name = "frmConciliaChequeCobranca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Porta As Integer, codigoSoma As String, strOrdem As String, StrOrdem2 As String

Private Sub Filtrar()
With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0 " & strOrdem
  .Refresh
End With
With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59# " & StrOrdem2
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59#"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and cobrando=0 and devolvido=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalPendente.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalPendente.Caption = Format(0, "Currency")
  End If
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from status"
  .Refresh
End With


End Sub

Private Sub ImprimeDados(ByVal Documento As String)
With dbClientes
  Y1 = Printer.CurrentY + 1
  .Recordset.FindFirst "cic='" & Documento & "'"
  If .Recordset.NoMatch = True Then
    .Recordset.FindFirst "cnpj='" & Documento & "'"
    If .Recordset.NoMatch = True Then Exit Sub
  End If
  Printer.FontSize = 8
  Printer.ForeColor = RGB(180, 180, 180)
  Printer.Line (0, Y1)-(190, Y1 + 33), , BF

  Printer.ForeColor = vbWhite
  Printer.Line (2, Y1 + 3)-(32, Y1 + 9), , BF
  Printer.Line (33, Y1 + 3)-(188, Y1 + 9), , BF
  Printer.Line (2, Y1 + 10)-(126, Y1 + 16), , BF
  Printer.Line (127, Y1 + 10)-(188, Y1 + 16), , BF
  Printer.Line (2, Y1 + 17)-(34, Y1 + 23), , BF
  Printer.Line (35, Y1 + 17)-(72, Y1 + 23), , BF
  Printer.Line (73, Y1 + 17)-(116, Y1 + 23), , BF
  Printer.Line (117, Y1 + 17)-(160, Y1 + 23), , BF
  Printer.Line (161, Y1 + 17)-(188, Y1 + 23), , BF
  Printer.Line (2, Y1 + 24)-(57, Y1 + 30), , BF
  Printer.Line (58, Y1 + 24)-(105, Y1 + 30), , BF
  Printer.Line (106, Y1 + 24)-(159, Y1 + 30), , BF
  Printer.Line (160, Y1 + 24)-(188, Y1 + 30), , BF
  
  Printer.FontName = "Arial"
  Printer.FontSize = 7
  Printer.ForeColor = vbBlack
  Printer.FillColor = vbBlack
  On Error Resume Next
  StrTemp = "Código"
  Printer.CurrentX = 3
  Printer.CurrentY = Y1 + 3
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 3
  StrTemp = .Recordset!codigochequecliente
  Printer.Print StrTemp
  
  StrTemp = "Nome"
  Printer.CurrentX = 34
  Printer.CurrentY = Y1 + 3
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 34
  StrTemp = .Recordset!Nome
  Printer.Print StrTemp
  
  StrTemp = "Endereço"
  Printer.CurrentX = 3
  Printer.CurrentY = Y1 + 10
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 3
  StrTemp = .Recordset!Endereco
  Printer.Print StrTemp
  
  StrTemp = "Bairro"
  Printer.CurrentX = 128
  Printer.CurrentY = Y1 + 10
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 128
  StrTemp = .Recordset!Codigo
  Printer.Print StrTemp
  
  StrTemp = "CEP"
  Printer.CurrentX = 3
  Printer.CurrentY = Y1 + 17
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 3
  StrTemp = .Recordset!CEP
  Printer.Print StrTemp
  
  StrTemp = "Telefone"
  Printer.CurrentX = 36
  Printer.CurrentY = Y1 + 17
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 36
  StrTemp = Format(.Recordset!Telefone, "(###)####-####")
  Printer.Print StrTemp
  
  
  StrTemp = "CIC"
  Printer.CurrentX = 74
  Printer.CurrentY = Y1 + 17
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 74
  StrTemp = Format(.Recordset!CIC, "##,###,###,###-##")
  Printer.Print StrTemp
  
  StrTemp = "RG"
  Printer.CurrentX = 118
  Printer.CurrentY = Y1 + 17
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 118
  StrTemp = Format(.Recordset!rg, "###,###,###,###-#")
  Printer.Print StrTemp
  
  StrTemp = "Emissão"
  Printer.CurrentX = 162
  Printer.CurrentY = Y1 + 17
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 162
  StrTemp = .Recordset!Origem & " - " & .Recordset!origem2
  Printer.Print StrTemp
  
  StrTemp = "CNPJ"
  Printer.CurrentX = 3
  Printer.CurrentY = Y1 + 24
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 3
  StrTemp = Format(.Recordset!CNPJ, "##,###,###/####-##")
  Printer.Print StrTemp
  
  StrTemp = "I.E."
  Printer.CurrentX = 59
  Printer.CurrentY = Y1 + 24
  Printer.Print StrTemp
  StrTemp = ""
  Printer.CurrentX = 59
  StrTemp = Format(.Recordset!ie, "###,###,###,###")
  Printer.Print StrTemp
  
  StrTemp = "Carro"
  Printer.CurrentX = 106
  Printer.CurrentY = Y1 + 24
  Printer.Print StrTemp
  StrTemp = ""
  If dbCarros.Recordset.EOF = False Then
    Printer.CurrentX = 106
    StrTemp = dbCarros.Recordset!Carro
    Printer.Print StrTemp
  End If
  
  StrTemp = "Placa"
  Printer.CurrentX = 161
  Printer.CurrentY = Y1 + 24
  Printer.Print StrTemp
  StrTemp = ""
  If dbCarros.Recordset.EOF = False Then
    Printer.CurrentX = 161
    StrTemp = dbCarros.Recordset!Placa
    Printer.Print StrTemp
  End If
  
  Printer.CurrentY = Y1 + 33
  Printer.FontSize = 10
End With
End Sub

Private Sub CabecaTodos(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Relção de Cheques Devolvidos"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp
StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

SubCabeca Largura
End Sub

Private Sub CabecaSoma(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters

Printer.FontBold = False
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Protocolo de Cheques Devolvidos Enviados ao Posto"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp
StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

SubCabeca2 Largura
End Sub

Private Sub SubCabeca(ByVal Largura As Double)
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = False
Printer.FontSize = 10

StrTemp = "Bom Para"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Comp"
Printer.CurrentX = 40 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Banco"
Printer.CurrentX = 65 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Agência"
Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cheque Nr."
Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1


End Sub

Private Sub SubCabeca2(ByVal Largura As Double)
Printer.CurrentY = Printer.CurrentY + 1

Printer.FontBold = False
Printer.FontSize = 10

StrTemp = "Bom Para"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Cobrando"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Comp"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Banco"
Printer.CurrentX = 70 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Agência"
Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cheque Nr."
Printer.CurrentX = 135 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cod.Cli."
Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1


End Sub

Private Sub ImprimeTodosCheques()
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbCheques
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0

  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 180
  Dia = Now
  
  
  CabecaTodos Largura, Dia
  
  Printer.FontSize = 10
  DiaAtual = .Recordset!datacheque
  Do While .Recordset.EOF = False
    If DiaAtual <> .Recordset!datacheque Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      StrTemp = "Sub-Total: " & Format(SubTotal, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      SubCabeca Largura
      SubTotal = 0
    End If
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.NewPage
      CabecaTodos Largura, Dia
    End If
    
    DiaAtual = .Recordset!datacheque
    StrTemp = .Recordset!datacheque
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!COMP
    Printer.CurrentX = 40 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 65 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    SubTotal = SubTotal + .Recordset!Valor
    Total = Total + .Recordset!Valor
    StrTemp = Format(.Recordset!Valor, "currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  StrTemp = "Sub-Total: " & Format(SubTotal, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
  
End With
NaoImprime:
End Sub

Private Sub ImprimeChequesSomadosSimples()
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontBold = False
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.DrawWidth = 2
  
  Largura = 180
  Dia = Now
  
  
  CabecaSoma Largura, Dia
  
  Printer.FontSize = 10
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 30 Then
      Printer.NewPage
      CabecaSoma Largura, Dia
    End If
    
    StrTemp = Format(.Recordset!datacheque, "Short date")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!datacobrando, "Short date")
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!COMP
    Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 70 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 135 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoCliente
    Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Total = Total + .Recordset!Valor
    StrTemp = Format(.Recordset!Valor, "currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.CurrentY = Printer.CurrentY + 2
  StrTemp = "Declaro estar em posse dos cheques descritos na lista acima."
  Printer.CurrentX = 0
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 5
  
  StrTemp = "Data: ______/______/____________"
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "________________________________________"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Assinatura"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "________________________________________"
  Printer.CurrentY = Printer.CurrentY + 5
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Nome"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.EndDoc
End With
NaoImprime:


End Sub

Private Sub ImprimeChequesSomados()
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.FontBold = False
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.DrawWidth = 2
  
  Largura = 180
  Dia = Now
  
  
  CabecaSoma Largura, Dia
  
  Printer.FontSize = 10
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 75 Then
      Printer.NewPage
      CabecaSoma Largura, Dia
    End If
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    
    StrTemp = Format(.Recordset!datacheque, "Short date")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!datacobrando, "Short date")
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!COMP
    Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 70 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 90 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Total = Total + .Recordset!Valor
    StrTemp = Format(.Recordset!Valor, "currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    If IsNull(.Recordset!CPF) = False Then
      ImprimeDados .Recordset!CPF
    End If
    
    Printer.Print ""
    StrTemp = "Data de contato com o cliente: ______/______/____________"
    Printer.CurrentX = 0
    Printer.Print StrTemp
    
    Printer.CurrentY = Printer.CurrentY + 5
    Printer.Line (3, Printer.CurrentY)-(Largura - 3, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 5
    Printer.Line (3, Printer.CurrentY)-(Largura - 3, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 5
    Printer.Line (3, Printer.CurrentY)-(Largura - 3, Printer.CurrentY)
    
    Printer.CurrentY = Printer.CurrentY + 1
    Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
    Printer.CurrentY = Printer.CurrentY + 1
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.CurrentY = Printer.CurrentY + 2
  StrTemp = "Declaro estar em posse dos cheques descritos na lista acima."
  Printer.CurrentX = 0
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 5
  
  StrTemp = "Data: ______/______/____________"
  Printer.CurrentX = 0
  Printer.Print StrTemp;
  
  StrTemp = "________________________________________"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Assinatura"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "________________________________________"
  Printer.CurrentY = Printer.CurrentY + 5
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  StrTemp = "Nome"
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub ProcuraCheque()
  Dim StrTemp As String
  lblData.Caption = ""
  lblValor.Caption = ""
  With dbCheques
    StrTemp = .Recordset.Sort
    .Refresh
    If MaskEdBox1(0).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "comp='" & MaskEdBox1(0).Text & "'"
      Else
        StrTemp = StrTemp & " and comp='" & MaskEdBox1(0).Text & "'"
      End If
    End If
    If MaskEdBox1(1).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "banco='" & MaskEdBox1(1).Text & "'"
      Else
        StrTemp = StrTemp & " and banco='" & MaskEdBox1(1).Text & "'"
      End If
    End If
    If MaskEdBox1(2).Text <> "    " Then
      If StrTemp = "" Then
        StrTemp = "agencia='" & MaskEdBox1(2).Text & "'"
      Else
        StrTemp = StrTemp & " and agencia='" & MaskEdBox1(2).Text & "'"
      End If
    End If
    If MaskEdBox1(3).Text <> "      - " Then
      If StrTemp = "" Then
        StrTemp = "conta='" & MaskEdBox1(3).Text & "'"
      Else
        StrTemp = StrTemp & " and conta='" & MaskEdBox1(3).Text & "'"
      End If
    End If
    If MaskEdBox1(4).Text <> "      " Then
      If StrTemp = "" Then
        StrTemp = "chequeNr='" & MaskEdBox1(4).Text & "'"
      Else
        StrTemp = StrTemp & " and chequeNr='" & MaskEdBox1(4).Text & "'"
      End If
    End If
    If StrTemp <> "" Then
      .Recordset.FindFirst StrTemp
      If .Recordset.NoMatch = False Then
        lblData.Caption = Format(.Recordset!datacheque, "short date")
        lblValor.Caption = Format(.Recordset!Valor, "Currency")
        Exit Sub
      End If
    End If
  End With
  StrTemp = ""
  With dbDepositar
    .Refresh
    If MaskEdBox1(0).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "comp='" & MaskEdBox1(0).Text & "'"
      Else
        StrTemp = StrTemp & " and comp='" & MaskEdBox1(0).Text & "'"
      End If
    End If
    If MaskEdBox1(1).Text <> "   " Then
      If StrTemp = "" Then
        StrTemp = "banco='" & MaskEdBox1(1).Text & "'"
      Else
        StrTemp = StrTemp & " and banco='" & MaskEdBox1(1).Text & "'"
      End If
    End If
    If MaskEdBox1(2).Text <> "    " Then
      If StrTemp = "" Then
        StrTemp = "agencia='" & MaskEdBox1(2).Text & "'"
      Else
        StrTemp = StrTemp & " and agencia='" & MaskEdBox1(2).Text & "'"
      End If
    End If
    If MaskEdBox1(3).Text <> "      - " Then
      If StrTemp = "" Then
        StrTemp = "conta='" & MaskEdBox1(3).Text & "'"
      Else
        StrTemp = StrTemp & " and conta='" & MaskEdBox1(3).Text & "'"
      End If
    End If
    If MaskEdBox1(4).Text <> "      " Then
      If StrTemp = "" Then
        StrTemp = "chequeNr='" & MaskEdBox1(4).Text & "'"
      Else
        StrTemp = StrTemp & " and chequeNr='" & MaskEdBox1(4).Text & "'"
      End If
    End If
    If StrTemp <> "" Then
      .Recordset.FindFirst StrTemp
      If .Recordset.NoMatch = False Then
        lblData.Caption = Format(.Recordset!datacheque, "short date")
        lblValor.Caption = Format(.Recordset!Valor, "Currency")
        Exit Sub
      End If
    End If
  End With
End Sub


Private Sub cboConta_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConta_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub CboConta_LostFocus()
Me.KeyPreview = True
With dbContas
  If cboConta.Text = "" Then Exit Sub
  .Refresh
  .Recordset.FindFirst "descri='" & cboConta.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboConta.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdConfirma_Click()
Dim Resposta As Integer, Diferenca As Currency
Dim strCheques As String, strDepositar As String
If dbDepositar.Recordset.RecordCount = 0 Then Exit Sub

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If dbDepositar.Recordset.EOF = True Then
  MsgBox "Selecione um cheque primeiro!"
  Exit Sub
End If

If DateDiff("d", Date, txtData.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber cheque com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -10 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber cheque com data anterior a 10 dias!"
    Exit Sub
  End If
End If

If cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Conta inválida!"
  cboConta.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido!"
  txtValor.SetFocus
  Exit Sub
End If

Resposta = MsgBox("Deseja confirmar o pagamento do cheque?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbDepositar
  Diferenca = CCur(txtValor.Text) - .Recordset!Valor
  If IsNull(dbDepositar.Recordset!CPF) = False Then
    dbClientes.Recordset.FindFirst "cic='" & dbDepositar.Recordset!CPF & "'"
  Else
    If IsNull(dbDepositar.Recordset!CNPJ) = False Then
      dbClientes.Recordset.FindFirst "cnpj='" & dbDepositar.Recordset!CNPJ & "'"
    Else
      dbClientes.Recordset.FindFirst "codigochequecliente=0"
    End If
  End If
  If dbClientes.Recordset.NoMatch = False Then
    dbClientes.Recordset.Edit
    dbClientes.Recordset!Devolvidos = dbClientes.Recordset!Devolvidos - 1
    dbClientes.Recordset!valordevolvido = dbClientes.Recordset!valordevolvido - .Recordset!Valor
    dbClientes.Recordset!saldopendente = dbClientes.Recordset!saldopendente - .Recordset!Valor
    '************************************************************************************************
    '************************************************************************************************
    'Desativado em 29/05/08 a reativação automática de clientes de cheque porque cliente que tem
    'cheque devolvido alinea 12 não será mais reativado
    '------------------------------------------------------------------------------------------------
    'If dbClientes.Recordset!Devolvidos = 0 Then
    '  dbClientes.Recordset!Posicao = True
    'End If
    '************************************************************************************************
    '************************************************************************************************
    dbClientes.Recordset.Update
  End If
  
  With dbStatus
    .Refresh
    .Recordset.Edit
    .Recordset!difcheques = .Recordset!difcheques + Diferenca
    .Recordset.Update
    .Refresh
  End With
  If dbContas.Recordset.EOF = False Then
    Call CboConta_LostFocus
    dbContas.Recordset.Edit
    dbContas.Recordset!Saldo = dbContas.Recordset!Saldo + CCur(txtValor.Text)
    dbContas.Recordset.Update
  End If
  With dbConciliaNova
    .Recordset.AddNew
    .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
    .Recordset!DataLanc = Now
    If dbContas.Recordset!temcpmf = True Then
      .Recordset!compensado = False
    Else
      .Recordset!compensado = True
      .Recordset!Data = Date
    End If
    .Recordset!Tipo = "Cheque Cobrança"
    .Recordset!Codigo = 999999996
    .Recordset!Descri = "Cobrança de Cheque Devolvido"
    .Recordset!NrDocumento = dbDepositar.Recordset!Banco & "/" & dbDepositar.Recordset!Agencia & "/" & dbDepositar.Recordset!Conta & "/" & dbDepositar.Recordset!chequenr
    .Recordset!Valor = CCur(txtValor.Text)
    .Recordset.Update
  End With
  .Recordset.Edit
  .Recordset!compensado = True
  .Recordset!cobrando = False
  .Recordset!devolvido = False
  .Recordset!datapgto = txtData.Value
  .Recordset!valorpgto = CCur(txtValor.Text)
  .Recordset!contacobrado = dbContas.Recordset!CodigoConta
  .Recordset!contadescricobrado = dbContas.Recordset!Descri
  .Recordset.Update
End With

codigoSoma = Str(CDbl(Now))

With dbCheques
  .Refresh
End With
With dbDepositar
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
dbDepositar.Recordset.Sort = strDepositar
txtValor.Text = ""
MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub cmdExibe_Click()
Filtrar
End Sub

Private Sub cmdImprime_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja imprimir relação de cheques em cobrança?" & Chr(vbKeyReturn) & "Sim - Imprime os cheques em cobrança," & Chr(vbKeyReturn) & "Não - Imprime cheques devolvidos em trânsito," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
Select Case Resposta
  Case vbYes
    Resposta = MsgBox("Deseja imprimir relatório detalhado de cheques em cobrança?" & Chr(vbKeyReturn) & "Sim - Imprime Relatório Detalhado," & Chr(vbKeyReturn) & "Não - Imprime listagem simples," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
    Select Case Resposta
      Case vbYes
        ImprimeChequesSomados
      Case vbNo
        ImprimeChequesSomadosSimples
      Case vbCancel
        Exit Sub
    End Select
  Case vbNo
    ImprimeTodosCheques
  Case vbCancel
    Exit Sub
End Select
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSomar_Click()

With dbCheques
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  For i = 0 To MaskEdBox1.Count - 1
    If Trim(MaskEdBox1(i).Text) = "" Then
      Resposta = MsgBox("Deseja incluir o cheque atual?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then
        Exit Sub
      Else
        Exit For
      End If
    End If
  Next i
  dbClientes.RecordSource = "select *from chequesclientes"
  dbClientes.Refresh
  If IsNull(.Recordset!CPF) = False Then
      dbClientes.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
  Else
    If IsNull(.Recordset!CNPJ) = False Then
      dbClientes.Recordset.FindFirst "cnpj='" & .Recordset!CNPJ & "'"
    Else
      dbClientes.Recordset.FindFirst "codigochequecliente=0"
    End If
  End If
  If dbClientes.Recordset.NoMatch = False Then
    dbClientes.Recordset.Edit
    dbClientes.Recordset!Devolvidos = dbClientes.Recordset!Devolvidos + 1
    dbClientes.Recordset!valordevolvido = dbClientes.Recordset!valordevolvido + .Recordset!Valor
    dbClientes.Recordset!Posicao = False
    dbClientes.Recordset.Update
  End If
  
  A = .Recordset!codigocheque
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then Exit Sub
  .Recordset.Edit
  .Recordset!cobrando = True
  .Recordset!datacobrando = Now
  .Recordset.Update
  .Refresh
End With
With dbDepositar
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and cobrando=0 and devolvido=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalPendente.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalPendente.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus
End Sub

Private Sub cmdSubtrair_Click()
With dbDepositar
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  
  If IsNull(.Recordset!CPF) = False Then
      dbClientes.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
  Else
    If IsNull(dbDepositar.Recordset!CNPJ) = False Then
      dbClientes.Recordset.FindFirst "cnpj='" & .Recordset!CNPJ & "'"
    Else
      dbClientes.Recordset.FindFirst "codigochequecliente=0"
    End If
  End If
  If dbClientes.Recordset.NoMatch = False Then
    dbClientes.Recordset.Edit
    dbClientes.Recordset!Devolvidos = dbClientes.Recordset!Devolvidos - 1
    dbClientes.Recordset!valordevolvido = dbClientes.Recordset!valordevolvido - .Recordset!Valor
    dbClientes.Recordset!Posicao = False
    dbClientes.Recordset.Update
  End If
  
  A = .Recordset!codigocheque
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then Exit Sub
  .Recordset.Edit
  .Recordset!cobrando = False
  .Recordset!datacobrando = Now
  .Recordset.Update
  .Refresh
  .Refresh
End With
With dbCheques
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and cobrando=0 and devolvido=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalPendente.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalPendente.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
With dbCheques
  If .RecordSource = "select *from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0 order by " & DBGrid1.Columns(ColIndex).DataField Then
    .RecordSource = "select *from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0 order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
  Else
    .RecordSource = "select *from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0 order by " & DBGrid1.Columns(ColIndex).DataField
  End If
  .Refresh
End With
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
With dbDepositar
  If .RecordSource = "select *from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59# order by " & DBGrid2.Columns(ColIndex).DataField & ", ChequeNr" Then
    .RecordSource = "select *from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59# order by " & DBGrid2.Columns(ColIndex).DataField & " desc, ChequeNr"
  Else
    .RecordSource = "select *from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59# order by " & DBGrid2.Columns(ColIndex).DataField & ", ChequeNr"
  End If
  .Refresh
End With
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
  Case Asc("+")
    KeyAscii = 0
    Call cmdSomar_Click
  Case Asc("-")
    KeyAscii = 0
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub Form_Load()
StrTemp = GetSetting(App.EXEName, "Base", "COM")

StrTemp2 = GetSetting(App.EXEName, "Base", "Baud", "9600")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "Paridade", "n")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "DataBit", "8")
StrTemp2 = StrTemp2 & "," & GetSetting(App.EXEName, "Base", "StopBit", "1")

MSComm1.Settings = StrTemp2

strOrdem = "order by chequenr"
StrOrdem2 = "order by CodigoCliente, ChequeNr"

With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from chequesclientes"
  .Refresh
End With
With dbClientesContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from chequescontas"
  .Refresh
End With
With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=0 and datacobrando Between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & " 00:00:00# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & " 23:59:59#" & StrOrdem2
  .Refresh
End With
With dbCarros
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from chequescarros"
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

If StrTemp <> "" Then
  If StrTemp <> "Sem" Then
    Porta = CInt(Right(StrTemp, 1))
  Else
    Porta = -1
  End If
End If

If Porta > 0 Then
  Timer1.Enabled = True
  MSComm1.CommPort = Porta
  Call Image1_DblClick
  On Error GoTo 0
End If
txtData.Value = Date
codigoSoma = Str(CDbl(Now))
txtDataIni.Value = Date
txtDataFim.Value = Date
Filtrar
Select Case Usuarios.Grupo.ChequeCobranca
  Case 1 'Somente leitura
    cmdSomar.Enabled = False
    cmdSubtrair.Enabled = False
    cmdConfirma.Enabled = False
    
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Image1_DblClick()
On Error Resume Next
With MSComm1
  If .PortOpen = True Then
    .PortOpen = False
  Else
    .PortOpen = True
  End If
  If .PortOpen = False Then
    Image1.Picture = LoadResPicture(102, vbResBitmap)
  Else
    Image1.Picture = LoadResPicture(101, vbResBitmap)
  End If
End With
End Sub

Private Sub MaskEdBox1_GotFocus(Index As Integer)
With MaskEdBox1(Index)
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub MaskEdBox1_LostFocus(Index As Integer)
  ProcuraCheque
End Sub



Private Sub Timer1_Timer()

If MSComm1.InBufferCount > 0 Then
  Timer1.Enabled = False

  'recebeu o codigo de barras armazena na variavel o codigo de barras
  CodBar = ""
  CodBar = MSComm1.Input
  If Len(CodBar) > 1 Then
    Do While Asc(Mid(CodBar, Len(CodBar) - 1, 1)) <> 3
      DoEvents
      CodBar = CodBar & MSComm1.Input
    Loop
    CodBar = Mid(CodBar, 1, Len(CodBar) - 1)
    CodBar = Converte(Trim(CodBar))
    If Len(CodBar) >= 33 Then
      'txtCodigo.Text = CodBar
      On Error Resume Next
      MaskEdBox1(0).Text = Mid(CodBar, 11, 3)
      MaskEdBox1(1).Text = Mid(CodBar, 2, 3)
      MaskEdBox1(2).Text = Mid(CodBar, 5, 4)
      MaskEdBox1(3).Text = Mid(CodBar, 26, 6) & "-" & Mid(CodBar, 32, 1)
      MaskEdBox1(4).Text = Mid(CodBar, 14, 6)
      
      ProcuraCheque
    End If
  End If
  Timer1.Enabled = True
End If

End Sub



Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub txtDataFim_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
If IsNumeric(txtValor.Text) = False Then Exit Sub
txtValor.Text = Format(txtValor.Text, "currency")
End Sub
