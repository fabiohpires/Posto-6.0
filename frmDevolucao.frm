VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaDevolucao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Devolução de Cheques"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmDevolucao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ConciliaNova"
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
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from estatus"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdEstornaCobranca 
      Caption         =   "Estorna Cobrança"
      Height          =   315
      Left            =   3480
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
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
   Begin VB.TextBox txtCMC7 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   3615
   End
   Begin VB.CommandButton cmdAntecipa 
      Caption         =   "Antecipa de Custódia"
      Height          =   495
      Left            =   6120
      TabIndex        =   34
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ChequesClientes"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbAlinea 
      Caption         =   "dbAlinea"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from alineas order by codigo"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdEstorna 
      Caption         =   "Estorna de Custódia"
      Height          =   495
      Left            =   8400
      TabIndex        =   33
      Top             =   2640
      Width           =   975
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from compensapendente where conciliado=0"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0 and codigosoma='1'"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbDepositar 
      Caption         =   "dbDepositar"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qSomaCheques 
      Caption         =   "qSomaCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   4200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qTotalPendente 
      Caption         =   "qTotalPendente"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where codigosoma='1' and compensado=0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   4920
      TabIndex        =   16
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37683
   End
   Begin VB.TextBox txtDescri 
      Height          =   285
      Left            =   2880
      TabIndex        =   14
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtLinha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar Devolução"
      Height          =   495
      Left            =   7320
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8520
      TabIndex        =   20
      Top             =   5160
      Width           =   855
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
      Left            =   6840
      TabIndex        =   18
      Top             =   2760
      Width           =   375
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
      Left            =   6360
      TabIndex        =   17
      Top             =   2760
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6360
      Top             =   840
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   120
      Picture         =   "frmDevolucao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "Imprimir"
      Top             =   5280
      Width           =   735
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
      Left            =   3840
      TabIndex        =   2
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
      Left            =   4440
      TabIndex        =   4
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
      Left            =   5040
      TabIndex        =   6
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
      Left            =   5760
      TabIndex        =   8
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
      Left            =   6720
      TabIndex        =   10
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
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmDevolucao.frx":0EC4
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmDevolucao.frx":0EDE
      TabIndex        =   31
      Top             =   3240
      Width           =   9255
   End
   Begin MSDBGrid.DBGrid grdCheques 
      Bindings        =   "frmDevolucao.frx":2625
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmDevolucao.frx":263D
      TabIndex        =   32
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label8 
      Caption         =   "Leitor de Código de barras:"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label70 
      Caption         =   "Leitura Automática"
      Height          =   255
      Left            =   1440
      TabIndex        =   30
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   1080
      Top             =   5400
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Data Devolução:"
      Height          =   195
      Left            =   4920
      TabIndex        =   15
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Descrição:"
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   2640
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alínea:"
      Height          =   195
      Left            =   2280
      TabIndex        =   11
      Top             =   2640
      Width           =   510
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      Caption         =   "Comp:"
      Height          =   195
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   450
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   2040
      Width           =   510
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      Height          =   195
      Left            =   5040
      TabIndex        =   5
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   5760
      TabIndex        =   7
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      Height          =   195
      Left            =   6720
      TabIndex        =   9
      Top             =   2040
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   1200
      TabIndex        =   27
      Top             =   2640
      Width           =   405
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   26
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Devolvido:"
      Height          =   195
      Left            =   3120
      TabIndex        =   25
      Top             =   5160
      Width           =   1170
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   24
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Total Depositado:"
      Height          =   255
      Left            =   8040
      TabIndex        =   23
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblTotalPendente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qTotalPendente"
      Height          =   255
      Left            =   8040
      TabIndex        =   22
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frmConciliaDevolucao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Porta As Integer, codigoSoma As String

Private Sub CabecaTodos(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Relção de Cheques Depositados"
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
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relação de Cheques Devolvidos"
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

Private Sub SubCabeca(ByVal Largura As Double)
Printer.CurrentY = Printer.CurrentY + 1

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

Private Sub ImprimeChequesSomados()
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbDepositar
  .RecordSource = "select *from cheques where somadevolucao='" & codigoSoma & "' order by datacheque, comp, agencia, banco, conta,chequenr, valor"
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 180
  Dia = Now
  
  
  CabecaSoma Largura, Dia
  
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
      CabecaSoma Largura, Dia
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
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub LimpaSoma()
  Dim Resposta As Integer
  
  With dbDepositar
    .RecordSource = "select *from cheques where devolvido=-1 and somadevolucao='Devolução'"
    .Refresh
    .Refresh
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      Resposta = MsgBox("Deseja remover os cheques da soma?", vbYesNo)
      If Resposta = vbNo Then Exit Sub
      Do While .Recordset.RecordCount <> 0
        .Recordset.Edit
        .Recordset!somadevolucao = "1"
        .Recordset!devolvido = False
        .Recordset!codigodevolucao = 0
        .Recordset!descridevolucao = " "
        .Recordset.Update
        .Refresh
        .Refresh
        .Refresh
      Loop
    End If
  End With
  Unload Me
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

Private Sub cmdAntecipa_Click()
Dim Resposta As Integer, Antecipado As Currency

If lblTotal.Caption = "" Then Exit Sub
If IsNumeric(lblTotal.Caption) = False Then Exit Sub
If CCur(lblTotal.Caption) = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Sub
End If
Antecipado = CCur(lblTotal.Caption)
Resposta = MsgBox("Deseja confirmar o estorno de custódia?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Selecione pelo menos um cheque!"
    Exit Sub
  End If
  Do While .Recordset.EOF = False
    With dbConcilia
      .Refresh
      .Recordset.FindFirst "data=#" & DataInglesa(dbDepositar.Recordset!datacheque) & "#"
      If .Recordset.NoMatch = True Then
        .Recordset.AddNew
        .Recordset!CodigoConta = dbDepositar.Recordset!CodigoConta
        .Recordset!Data = dbDepositar.Recordset!datacheque
        .Recordset!Descri = "Custódia de cheques!"
        .Recordset!Valor = -dbDepositar.Recordset!Valor
        .Recordset!Conta = dbDepositar.Recordset!contadescri
        .Recordset.Update
      Else
        .Recordset.Edit
        .Recordset!Valor = .Recordset!Valor - dbDepositar.Recordset!Valor
        .Recordset.Update
      End If
    End With
    dbContas.Refresh
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset!CodigoConta
    If dbContas.Recordset.EOF = False Then
      dbContas.Recordset.Edit
      dbContas.Recordset!Saldo = dbContas.Recordset!Saldo + .Recordset!Valor
      dbContas.Recordset.Update
    End If
    
    If IsNull(.Recordset!CPF) = False Then
      dbClientes.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
    Else
      If IsNull(dbDepositar.Recordset!CNPJ) = False Then
        dbClientes.Recordset.FindFirst "cnpj='" & .Recordset!CNPJ & "'"
      Else
        dbClientes.Recordset.FindFirst "codigochequecliente=0"
      End If
    End If
    
    .Recordset.Edit
    .Recordset!codigoSoma = "1"
    .Recordset!somadevolucao = "1"
    .Recordset.Update
    .Recordset.MoveNext
  Loop
End With
With dbConcilia
  .RecordSource = "select *from concilianova"
  .Refresh
  .Recordset.AddNew
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  .Recordset!compensado = True
  .Recordset!Data = txtData.Value
  .Recordset!Tipo = "Antecipação"
  .Recordset!Codigo = 999999994
  .Recordset!Descri = "Custódia Antecipada"
  .Recordset!NrDocumento = codigoSoma
  .Recordset!Valor = Antecipado
  .Recordset.Update
  .RecordSource = "select *from compensapendente where conciliado=0"
  .Refresh
End With
codigoSoma = GeraCodigo() ' Str(CDbl(Now))

With dbCheques
  .Refresh
End With
With dbDepositar
  .RecordSource = "select *from cheques where somadevolucao='Devolução'"
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub cmdConfirma_Click()
Dim Resposta As Integer

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With


If lblTotal.Caption = "" Then Exit Sub
If IsNumeric(lblTotal.Caption) = False Then Exit Sub
If CCur(lblTotal.Caption) = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Sub
End If
Resposta = MsgBox("Deseja confirmar a relação de cheques devolvidos?", vbYesNo)
If Resposta = vbNo Then Exit Sub

Resposta = MsgBox("Deseja imprimir a relação de cheques devolvidos?", vbYesNo)
If Resposta = vbYes Then
  ImprimeChequesSomados
End If


With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Selecione pelo menos um cheque!"
    Exit Sub
  End If
  Do While .Recordset.EOF = False
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
      dbClientes.Recordset!Depositados = dbClientes.Recordset!Depositados - 1
      dbClientes.Recordset!valordepositado = dbClientes.Recordset!valordepositado - .Recordset!Valor
      dbClientes.Recordset!Devolvidos = dbClientes.Recordset!Devolvidos + 1
      dbClientes.Recordset!valordevolvido = dbClientes.Recordset!valordevolvido + .Recordset!Valor
      dbClientes.Recordset!Posicao = False
      dbClientes.Recordset!datadesativado = Now
      dbClientes.Recordset.Update
    End If
    
    dbContas.Refresh
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset!CodigoConta
    If dbContas.Recordset.EOF = False Then
      dbContas.Recordset.Edit
      dbContas.Recordset!Saldo = dbContas.Recordset!Saldo - .Recordset!Valor
      dbContas.Recordset.Update
    End If
    With dbConcilia
      .RecordSource = "select *from concilianova"
      .Refresh
      .Recordset.AddNew
      .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
      .Recordset!DataLanc = Now
      .Recordset!compensado = True
      .Recordset!Data = dbDepositar.Recordset!bancodevolucao
      .Recordset!Tipo = "Devolução"
      .Recordset!Codigo = 999999994
      .Recordset!Descri = "Cheque Dep. Devolvido"
      .Recordset!NrDocumento = dbDepositar.Recordset!Banco & "/" & dbDepositar.Recordset!Agencia & "/" & dbDepositar.Recordset!Conta & "/" & dbDepositar.Recordset!chequenr
      .Recordset!Valor = -dbDepositar.Recordset!Valor
      .Recordset.Update
      .RecordSource = "select *from compensapendente where conciliado=0"
      .Refresh
    End With

    .Recordset.Edit
    .Recordset!codigoSoma = "1"
    .Recordset!somadevolucao = "1"
    .Recordset!compensado = False
    .Recordset!Custodia = False
    .Recordset.Update
    .Recordset.MoveNext
  Loop
End With

codigoSoma = GeraCodigo() ' Str(CDbl(Now))

With dbCheques
  .Refresh
End With
With dbDepositar
  .RecordSource = "select *from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub cmdEstorna_Click()
Dim Resposta As Integer

If lblTotal.Caption = "" Then Exit Sub
If IsNumeric(lblTotal.Caption) = False Then Exit Sub
If CCur(lblTotal.Caption) = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Sub
End If
Resposta = MsgBox("Deseja confirmar o estorno de custódia?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Selecione pelo menos um cheque!"
    Exit Sub
  End If
  Do While .Recordset.EOF = False
    With dbConcilia
      .Refresh
      .Recordset.FindFirst "data=#" & DataInglesa(dbDepositar.Recordset!datacheque) & "#"
      If .Recordset.NoMatch = True Then
        .Recordset.AddNew
        .Recordset!CodigoConta = dbDepositar.Recordset!CodigoConta
        .Recordset!Data = dbDepositar.Recordset!datacheque
        .Recordset!Descri = "Estorno de Custódia!"
        .Recordset!Valor = -dbDepositar.Recordset!Valor
        .Recordset!Conta = dbDepositar.Recordset!contadescri
        .Recordset.Update
      Else
        .Recordset.Edit
        .Recordset!Valor = .Recordset!Valor - dbDepositar.Recordset!Valor
        .Recordset.Update
      End If
    End With
    dbContas.Refresh
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset!CodigoConta
    
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
      dbClientes.Recordset!Depositados = dbClientes.Recordset!Depositados - 1
      dbClientes.Recordset!valordepositado = dbClientes.Recordset!valordepositado - .Recordset!Valor
      dbClientes.Recordset.Update
    End If
    
    .Recordset.Edit
    .Recordset!codigoSoma = "1"
    .Recordset!somadevolucao = "1"
    .Recordset!compensado = False
    .Recordset!devolvido = False
    .Recordset!Custodia = False
    .Recordset.Update
    .Recordset.MoveNext
  Loop
End With

codigoSoma = GeraCodigo() ' Str(CDbl(Now))

With dbCheques
  .Refresh
End With
With dbDepositar
  .RecordSource = "select *from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub cmdEstornaCobranca_Click()
Dim Resposta As Integer, Diferenca As Currency
Dim strCheques As String, strDepositar As String
If dbCheques.Recordset.RecordCount = 0 Then Exit Sub

If dbCheques.Recordset.EOF = True Then
  MsgBox "Selecione um cheque primeiro!"
  Exit Sub
End If

If dbCheques.Recordset!valorpgto = 0 Then
  MsgBox "Este cheque não foi cobrado!"
  Exit Sub
End If

Resposta = MsgBox("Deseja estornar o pagamento do cheque?", vbYesNo)
If Resposta = vbNo Then Exit Sub

With dbCheques
  Diferenca = .Recordset!valorpgto - .Recordset!Valor
  If IsNull(dbDepositar.Recordset!CPF) = False Then
    dbClientes.Recordset.FindFirst "cic='" & dbCheques.Recordset!CPF & "'"
  Else
    If IsNull(dbDepositar.Recordset!CNPJ) = False Then
      dbClientes.Recordset.FindFirst "cnpj='" & dbCheques.Recordset!CNPJ & "'"
    Else
      dbClientes.Recordset.FindFirst "codigochequecliente=0"
    End If
  End If
  If dbClientes.Recordset.NoMatch = False Then
    dbClientes.Recordset.Edit
    dbClientes.Recordset!Devolvidos = dbClientes.Recordset!Devolvidos + 1
    dbClientes.Recordset!valordevolvido = dbClientes.Recordset!valordevolvido + .Recordset!Valor
    dbClientes.Recordset!saldopendente = dbClientes.Recordset!saldopendente + .Recordset!Valor
    dbClientes.Recordset!Posicao = False
    dbClientes.Recordset.Update
  End If
  
  
  With dbStatus
    .Refresh
    .Recordset.Edit
    .Recordset!difcheques = .Recordset!difcheques - Diferenca
    .Recordset.Update
    .Refresh
  End With
  If dbContas.Recordset.EOF = False Then
    dbContas.Recordset.FindFirst "codigoconta=" & dbCheques.Recordset!contacobrado
    dbContas.Recordset.Edit
    dbContas.Recordset!Saldo = dbContas.Recordset!Saldo - dbCheques.Recordset!valorpgto
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
    .Recordset!Descri = "Estorno de Cobrança de Ch. Dev."
    .Recordset!NrDocumento = dbCheques.Recordset!Banco & "/" & dbCheques.Recordset!Agencia & "/" & dbCheques.Recordset!Conta & "/" & dbCheques.Recordset!chequenr
    .Recordset!Valor = -dbCheques.Recordset!valorpgto
    .Recordset.Update
  End With
  .Recordset.Edit
  .Recordset!compensado = False
  .Recordset!cobrando = True
  .Recordset!devolvido = True
  .Recordset!valorpgto = 0
  .Recordset!contacobrado = 0
  .Recordset!contadescricobrado = " "
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
MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus


End Sub

Private Sub cmdImprime_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja imprimir relação de cheques devolvidos?" & Chr(vbKeyReturn) & "Sim - Imprime os cheques devolvidos," & Chr(vbKeyReturn) & "Não - Imprime todos os cheques," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
Select Case Resposta
  Case vbYes
    ImprimeChequesSomados
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

If DateDiff("d", Date, txtData.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode devolver com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -15 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode devolver com data anterior a 15 dias!"
    Exit Sub
  End If
End If

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
  If IsNumeric(txtLinha.Text) = False Then
    MsgBox "Informe o número da linha de devolução!"
    txtLinha.SetFocus
    Exit Sub
  End If
  If txtDescri.Text = "" Then txtDescri.Text = " "
  
  
  
  A = .Recordset!codigocheque
  StrTemp = .Recordset.Sort
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then Exit Sub
  .Recordset.Edit
  .Recordset!somadevolucao = "Devolução"
  .Recordset!datadevolucao = Now
  .Recordset!bancodevolucao = txtData.Value
  .Recordset!devolvido = True
  .Recordset!codigodevolucao = CInt(txtLinha.Text)
  .Recordset!descridevolucao = txtDescri.Text
  .Recordset.Update
  A = .Recordset!codigocheque
  .Refresh
End With
With dbDepositar
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='Devolução'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

'MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
'txtLinha.Text = ""
'txtDescri.Text = ""
txtCMC7.Text = ""

txtCMC7.SetFocus
End Sub

Private Sub cmdSubtrair_Click()
With dbDepositar
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  A = .Recordset!codigocheque
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then Exit Sub
  .Recordset.Edit
  .Recordset!somadevolucao = "1"
  .Recordset!devolvido = False
  .Recordset!codigodevolucao = 0
  .Recordset!descridevolucao = " "
  .Recordset.Update
  .Refresh
End With
With dbCheques
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='Devolução'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
If dbDepositar.RecordSource = "select *from cheques where somadevolucao='Devolução' order by " & grdCheques.Columns(ColIndex).DataField Then
  dbDepositar.RecordSource = "select *from cheques where somadevolucao='Devolução' order by " & grdCheques.Columns(ColIndex).DataField & " desc"
Else
  dbDepositar.RecordSource = "select *from cheques where somadevolucao='Devolução' order by " & grdCheques.Columns(ColIndex).DataField
End If
dbDepositar.Refresh
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
End If
txtData.Value = Date
codigoSoma = GeraCodigo()

With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where compensado=-1 or custodia=-1 and somadevolucao='1' order by chequenr, valor"
  .Refresh
End With
With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where somadevolucao='Devolução'"
  .Refresh
  .Refresh
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='Devolução'"
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
  .RecordSource = "select sum(valor) as total from cheques where somadevolucao='1' and compensado=-1"
  .Refresh
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbAlinea
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With
With dbStatus
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from status"
  .Refresh
End With

Select Case Usuarios.Grupo.ChequeDevolucao
  Case 1 'Somente leitura
    cmdSomar.Enabled = False
    cmdSubtrair.Enabled = False
    cmdConfirma.Enabled = False
    cmdEstorna.Enabled = False
    cmdAntecipa.Enabled = False
  Case 2 'Liberado
    
End Select
If Usuarios.Nome = "Usuário Master" Then
  cmdEstornaCobranca.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
LimpaSoma
End Sub

Private Sub grdCheques_HeadClick(ByVal ColIndex As Integer)
If dbCheques.RecordSource = "select *from cheques where compensado=-1 and somadevolucao='1' order by " & grdCheques.Columns(ColIndex).DataField Then
  dbCheques.RecordSource = "select *from cheques where compensado=-1 and somadevolucao='1' order by " & grdCheques.Columns(ColIndex).DataField & " desc"
Else
  dbCheques.RecordSource = "select *from cheques where compensado=-1 and somadevolucao='1' order by " & grdCheques.Columns(ColIndex).DataField
End If
dbCheques.Refresh
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
      txtLinha.SetFocus
    End If
  End If
  Timer1.Enabled = True
End If

End Sub


Private Sub txtCMC7_Change()
Dim Cheque As DadosCheque
Dim Cheque2 As CMC7

Cheque = ConverteCMC7(txtCMC7.Text)

If Cheque.COMP = "" Then Exit Sub

Cheque2.CMC7 = txtCMC7.Text
Cheque2 = CMC7Define(Cheque2)
If Cheque2.Validado = True Then
  With Cheque
    MaskEdBox1(0).Text = .COMP
    MaskEdBox1(1).Text = .Banco
    MaskEdBox1(2).Text = .Agencia
    MaskEdBox1(3).Text = .Conta
    MaskEdBox1(4).Text = .Cheque
    CodBar = txtCMC7.Text
    MaskEdBox1(0).SetFocus
    cmdSomar.SetFocus
  End With
End If

End Sub

Private Sub txtLinha_LostFocus()
With dbAlinea
  If txtLinha.Text <> "" Then
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.FindFirst "codigo=" & txtLinha.Text
      If .Recordset.NoMatch = False Then
        txtDescri.Text = .Recordset!Descricao
        Exit Sub
      End If
    End If
  End If
End With
Select Case txtLinha.Text
  Case "11"
    txtDescri.Text = "Cheques sem fundos 1ª Ap."
  Case "12"
    txtDescri.Text = "Cheques sem fundos 2ª Ap."
  Case "00"
    txtDescri.Text = "Sem motivo informado"
  Case "20"
    txtDescri.Text = "Folha de cheque cancelada"
  Case "21"
    txtDescri.Text = "Cheque sustado"
  Case "25"
    txtDescri.Text = "Cancelamento Talonário"
  Case "48"
    txtDescri.Text = "Cheque sem nominal"
  Case "28"
    txtDescri.Text = "Contra ordem ou oposição"
  Case "29"
    txtDescri.Text = "Talão bloqueado"
  Case "49"
    txtDescri.Text = "Cheque sem fundos 3ª Ap."
End Select
End Sub
