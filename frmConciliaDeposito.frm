VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaDeposito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depositar Cheques"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmConciliaDeposito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGeraArquivo 
      Caption         =   "Gerar Arquivo"
      Height          =   375
      Left            =   2400
      TabIndex        =   37
      Top             =   6000
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7440
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   600
      Top             =   2280
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
      Top             =   360
      Width           =   3615
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ChequesClientes"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from concilianova"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmConciliaDeposito.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   3720
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDBGrid.DBGrid grdCheques 
      Bindings        =   "frmConciliaDeposito.frx":0459
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmConciliaDeposito.frx":0471
      TabIndex        =   32
      Top             =   1440
      Width           =   9255
   End
   Begin VB.Data qTotalPendente 
      Caption         =   "qTotalPendente"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where codigosoma='1' and compensado=0 and protesto=0"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qSomaCheques 
      Caption         =   "qSomaCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from compensapendente"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbDepositar 
      Caption         =   "dbDepositar"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0"
      Top             =   5040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "\\Posto01\Rede\Dados\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0 and codigosoma='1' and protesto=0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdCustódia 
      Caption         =   "Custódia"
      Height          =   375
      Left            =   6240
      TabIndex        =   31
      Top             =   3600
      Width           =   855
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   3600
      TabIndex        =   29
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37767
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Ö"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   28
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox txtDocumemto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6120
      TabIndex        =   12
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8640
      Picture         =   "frmConciliaDeposito.frx":1D6B
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "Imprimir"
      Top             =   3480
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6480
      Top             =   2520
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5880
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
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
      Left            =   8520
      TabIndex        =   13
      Top             =   240
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
      Left            =   9000
      TabIndex        =   14
      Top             =   240
      Width           =   375
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Depósito"
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   3600
      Width           =   855
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
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
      TabIndex        =   4
      Top             =   960
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
      TabIndex        =   6
      Top             =   960
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
      TabIndex        =   8
      Top             =   960
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
      TabIndex        =   10
      Top             =   960
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
      Bindings        =   "frmConciliaDeposito.frx":27ED
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmConciliaDeposito.frx":2807
      TabIndex        =   33
      Top             =   4080
      Width           =   9255
   End
   Begin VB.Label Label8 
      Caption         =   "Leitor de Código de barras:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblCheques 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   3600
      TabIndex        =   30
      Top             =   3480
      Width           =   390
   End
   Begin VB.Label Label4 
      Caption         =   "CPF/CNPJ:"
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label70 
      Caption         =   "Leitura Automática"
      Height          =   255
      Left            =   480
      TabIndex        =   27
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Top             =   6000
      Width           =   255
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
      Height          =   315
      Left            =   7320
      TabIndex        =   25
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Total Pendente:"
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   465
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2280
      TabIndex        =   22
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   2280
      TabIndex        =   21
      Top             =   3480
      Width           =   405
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   4920
      TabIndex        =   17
      Top             =   120
      Width           =   405
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   3840
      TabIndex        =   15
      Top             =   120
      Width           =   390
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   2040
      TabIndex        =   7
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      Height          =   195
      Left            =   1320
      TabIndex        =   5
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      Caption         =   "Comp:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   450
   End
End
Attribute VB_Name = "frmConciliaDeposito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Porta As Integer, codigoSoma As String, ColunaQuebra As Integer
Dim Lote As String

Private Sub CabecaBordero(ByVal Largura As Double, Dia As Date, ByVal NumeroBordero As Double)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Relção de Cheques para Custódia"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp
StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = "Borderô: " & Format(NumeroBordero, "0000000")
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Agência: " & dbContas.Recordset!Agencia & "    Conta Corrente: " & dbContas.Recordset!cc & "     Empresa: " & dbContas.Recordset!CodigoEmpresa & "    Filial: " & dbContas.Recordset!CodigoFilial
Printer.CurrentX = 0
Printer.Print StrTemp

StrTemp = "Quantidade: " & dbDepositar.Recordset.RecordCount & "    Total: " & lblTotal.Caption
Printer.CurrentX = 0
Printer.Print StrTemp

SubCabeca Largura
End Sub

Public Sub BorderoBradesco(ByVal NumeroBordero As Double)
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbDepositar
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 180
  Dia = Now
  
  
  CabecaBordero Largura, Dia, NumeroBordero
  
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
      CabecaBordero Largura, Dia, NumeroBordero
    End If
    
    DiaAtual = .Recordset!datacheque
    StrTemp = .Recordset!datacheque
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!COMP
    Printer.CurrentX = 35 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 75 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoCliente
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
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

Public Function CustodiaBB() As Boolean
Dim nArq As Integer
nArq = FreeFile()
CustodiaBB = False
Open "Custodia.txt" For Output As nArq
With dbDepositar
  .Recordset.MoveFirst
  
  Do While .Recordset.EOF = False
    StrTemp = String(240, "0")
    Mid(StrTemp, 18, 3) = .Recordset!COMP
    Mid(StrTemp, 21, 3) = .Recordset!Banco
    Mid(StrTemp, 25, 4) = .Recordset!Agencia
    StrTemp2 = Right(String(12, "0") & .Recordset!Conta, 12)
    A = InStr(1, StrTemp2, "-")
    If A <> 0 Then
      StrTemp2 = "0" & Mid(StrTemp2, 1, A - 1) & Mid(StrTemp2, A + 1)
    End If
    Mid(StrTemp, 30, 12) = StrTemp2
    Mid(StrTemp, 43, 6) = .Recordset!chequenr
    StrTemp2 = Right(String(15, "0") & Format(.Recordset!Valor, "0.00"), 15)
    A = InStr(1, StrTemp2, ",")
    If A <> 0 Then
      StrTemp2 = "0" & Mid(StrTemp2, 1, A - 1) & Mid(StrTemp2, A + 1)
    End If
    Mid(StrTemp, 66, 15) = StrTemp2
    StrTemp2 = Format(Day(.Recordset!datacheque), "00")
    Mid(StrTemp, 81, 2) = StrTemp2
    StrTemp2 = Format(Month(.Recordset!datacheque), "00")
    Mid(StrTemp, 83, 2) = StrTemp2
    StrTemp2 = Format(Year(.Recordset!datacheque), "0000")
    Mid(StrTemp, 85, 4) = StrTemp2
    
    Print nArq, StrTemp
    .Recordset.MoveNext
  Loop
  Close nArq
End With
CustodiaBB = True
End Function

Public Function CustodiaBradesco() As Boolean
Dim db As New ADODB.Connection, dbConfig As New ADODB.Recordset
Dim nArq As Integer, StrArquivo As String, StrTemp As String
Dim DataArquivo As Date, PrimeiraLinha As String
Dim StrConta As String, strDigitoConta As String, strAgencia As String
Dim PoloOrigem As String, Bordero As String
Dim CodigoEmpresa As Double, CodigoFilial As String, strTempValor As String
Dim DataLimite As Date


nArq = FreeFile()
CustodiaBradesco = False

If dbDepositar.Recordset.RecordCount = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Function
End If
With dbContas
  If .Recordset.RecordCount = 0 Then Exit Function
  If cboConta.Text = "" Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  Call CboConta_LostFocus
  If cboConta.Text <> .Recordset!Descri Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  
  If IsNull(.Recordset!cc) = True Then
    MsgBox "Complete o cadastro da conta atual!"
    Exit Function
  End If
  If IsNull(.Recordset!CodigoEmpresa) = True Then
    MsgBox "Complete o cadastro da conta atual!"
    Exit Function
  Else
    CodigoEmpresa = .Recordset!CodigoEmpresa
  End If
  If IsNull(.Recordset!CodigoFilial) = True Then
    MsgBox "Complete o cadastro da conta atual!"
    Exit Function
  Else
    CodigoFilial = .Recordset!CodigoFilial
  End If
  If IsNull(.Recordset!custodialote) = False Then
    If .Recordset!custodialote > 0 Then
      Lote = .Recordset!custodialote
    Else
      Lote = 1
    End If
  Else
    Lote = 1
    dbContas.Recordset.Edit
    dbContas.Recordset!custodialote = 1
    dbContas.Recordset.Update
  End If
  StrConta = .Recordset!cc
  strAgencia = .Recordset!Agencia
End With

A = InStr(1, StrConta, "-")
If A = 0 Then
  MsgBox "Número da conta deve ser cadastrado com o dígito verificador! Ex: '123.456-0'"
  Exit Function
End If

With dbDepositar
  DataLimite = DateAdd("d", 1, Date)
  .Recordset.FindFirst "datacheque<=#" & DataInglesa(DataLimite) & "#"
  If .Recordset.NoMatch = False Then
    MsgBox "Existe cheque com data inferior ao limite para ser custodiado!"
    Exit Function
  End If
End With

strDigitoConta = Mid(StrConta, A + 1)
StrConta = Mid(StrConta, 1, A - 1)

StrConta = RemoveString(StrConta)

db.Open CaminhoADO
dbConfig.Open "Select *from config", db, adOpenKeyset, adLockOptimistic

StrArquivo = Format(Now, "YYYYmmDD") & "-" & Format(Now, "HHNN") & ".DAT"

If dbConfig.RecordCount <> 0 Then
  If IsNull(dbConfig!localcustodia) = False Then
    If Right(dbConfig!localcustodia, 1) <> "\" Then
      StrArquivo = dbConfig!localcustodia & "\" & StrArquivo
    Else
      StrArquivo = dbConfig!localcustodia & StrArquivo
    End If
  End If
End If

Open StrArquivo For Output As #nArq

'*******************************************************************************************************
'*******************************************************************************************************
PoloOrigem = "4470"
'*******************************************************************************************************
'*******************************************************************************************************

PrimeiraLinha = Space(250)
Mid(PrimeiraLinha, 1) = "0"
Mid(PrimeiraLinha, 2) = Format(Date, "yyyymmdd")
Mid(PrimeiraLinha, 10) = "CUSTODIA"
Mid(PrimeiraLinha, 18) = "00000"
Mid(PrimeiraLinha, 23) = Format(CodigoEmpresa, "000000")
Mid(PrimeiraLinha, 29) = Format(Time, "hhnnss")
Mid(PrimeiraLinha, 35) = "237"
Mid(PrimeiraLinha, 38) = Format(strAgencia, "00000")
Mid(PrimeiraLinha, 43) = Format(StrConta, "0000000000000")
Mid(PrimeiraLinha, 56) = strDigitoConta
Mid(PrimeiraLinha, 57) = Space(1)
Mid(PrimeiraLinha, 58) = Format(CodigoFilial, "00000")
Mid(PrimeiraLinha, 63) = Space(118)
Mid(PrimeiraLinha, 181) = "0000000001"
Mid(PrimeiraLinha, 191) = Space(60)

Print #nArq, PrimeiraLinha

With dbDepositar
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    StrTemp = Space(250)
    If IsNull(.Recordset!CMC7) = True Then
      MsgBox "Existe cheque sem CMC7 cadastrado!"
      Close nArq
      Exit Function
    End If
    
    StrTemp2 = .Recordset!CMC7
    Mid(StrTemp, 1) = "1"
    Mid(StrTemp, 2) = Mid(StrTemp2, 11, 3) 'Comp
    Mid(StrTemp, 5) = Mid(StrTemp2, 2, 3) 'banco
    Mid(StrTemp, 8) = Mid(StrTemp2, 5, 4) 'agencia
    Mid(StrTemp, 12) = Mid(StrTemp2, 22, 1) 'c1
    Mid(StrTemp, 13) = Mid(StrTemp2, 26, 7) 'conta
    Mid(StrTemp, 25) = Mid(StrTemp2, 9, 1) 'c2
    Mid(StrTemp, 26) = Mid(StrTemp2, 14, 6) 'nr cheque
    Mid(StrTemp, 32) = Mid(StrTemp2, 33, 1) 'c3
    strTempValor = Format(.Recordset!Valor, "000000000000000.00")
    strTempValor = RemoveString(strTempValor)
    Mid(StrTemp, 33) = strTempValor
    Mid(StrTemp, 50) = Mid(StrTemp2, 20, 1) 'tipo do cheque
    Mid(StrTemp, 12) = Format(Date, "yyyymmdd")
    Mid(StrTemp, 12) = Format(.Recordset!datacheque, "yyyymmdd")
    If Fu_consistir_CgcCpf(.Recordset!CPF) = True Then '??????????????????????????
      StrTemp2 = .Recordset!CPF
      StrTemp2 = RemoveString(StrTemp2)
      Mid(StrTemp, 67) = Format(StrTemp2, "0000000000000000")
      Mid(StrTemp, 67) = "0004"
    Else
      Mid(StrTemp, 67) = "0000000000000000"
      Mid(StrTemp, 67) = "0006"
    End If
    Mid(StrTemp, 83) = "000"
    Mid(StrTemp, 86) = String(20, "0")
    Bordero = Format(.Recordset!codigocheque, "0000000")
    Mid(StrTemp, 106) = Format(Bordero, "0000000") ' número do bordero
    Mid(StrTemp, 120) = Space(25) ' para uso do cliente
    Mid(StrTemp, 145) = Space(36)
    Mid(StrTemp, 181) = Format(.Recordset.AbsolutePosition + 1, "0000000000")
    Mid(StrTemp, 145) = Space(60)
    
    Print #nArq, StrTemp
    .Recordset.MoveNext
  Loop
  
End With

PrimeiraLinha = Space(250)
Mid(PrimeiraLinha, 1) = "9"
Mid(PrimeiraLinha, 2) = Format(Date, "yyyymmdd")
Mid(PrimeiraLinha, 10) = "CUSTODIA"
Mid(PrimeiraLinha, 18) = Format(dbDepositar.Recordset.RecordCount, "000000000")
strTempValor = RemoveString(Format(lblTotal.Caption, "000000000000000.00"))
Mid(PrimeiraLinha, 27) = strTempValor
Mid(PrimeiraLinha, 44) = Space(137)
Mid(PrimeiraLinha, 181) = Format(Lote, "0000000000")
Mid(PrimeiraLinha, 191) = Space(60)

Print #nArq, PrimeiraLinha

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Function
On Error GoTo 0


BorderoBradesco Lote
BorderoBradesco Lote

CustodiaBradesco = True

NaoImprime:

Close #nArq

End Function



Public Function DepositoBradesco() As Boolean
Dim nArq As Integer, StrArquivo As String, StrTemp As String
Dim DataArquivo As Date, PrimeiraLinha As String

nArq = FreeFile()
DepositoBradesco = False

If dbDepositar.Recordset.RecordCount = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Function
End If
With dbContas
  If .Recordset.RecordCount = 0 Then Exit Function
  If cboConta.Text = "" Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  Call CboConta_LostFocus
  If cboConta.Text <> .Recordset!Descri Then
    MsgBox "Selecione uma conta!"
    Exit Function
  End If
  
  PrimeiraLinha = Space(46)
  Mid(PrimeiraLinha, 1, 4) = Format(.Recordset!Agencia, "0000")
  Mid(PrimeiraLinha, 5, 18) = Format(lblTotal.Caption, "000000000000000.00")
  
End With

With dbContas
  If IsNull(.Recordset!cc) = True Then
    MsgBox "Complete o cadastro da conta atual!"
    Exit Function
  End If
  StrArquivo = .Recordset!cc
End With

A = InStr(1, StrArquivo, "-")
If A = 0 Then
  MsgBox "Número da conta deve ser cadastrado com o dígito verificador! Ex: '123.456-0'"
  Exit Function
End If

StrArquivo = Mid(StrArquivo, 1, A - 1)
A = 0
A = InStr(1, StrArquivo, ".")


Do While A <> 0
  StrArquivo = Mid(StrArquivo, 1, A - 1) & Mid(StrArquivo, A + 1)
  A = 0
  A = InStr(1, StrArquivo, ".")
Loop

A = 1

StrTemp = "0000000"
Mid(StrTemp, 8 - Len(StrArquivo)) = StrArquivo
StrArquivo = StrTemp

StrTemp = "c:\" & StrArquivo & ".001"
If Dir(StrTemp) <> "" Then
  DataArquivo = FileDateTime(StrTemp)
  DataArquivo = CDate(Format(DataArquivo, "short date"))
  
  If DataArquivo <> Date Then
    Kill StrTemp
  End If
  Do While Dir(StrTemp) <> ""
    A = A + 1
    StrTemp = "c:\" & StrArquivo & "." & Format(A, "000")
  Loop
  StrArquivo = StrTemp
Else
  StrArquivo = StrTemp
End If

Open StrArquivo For Output As #nArq

Print #nArq, PrimeiraLinha

With dbDepositar
  
  .Recordset.MoveFirst
  
  Do While .Recordset.EOF = False
    StrTemp = Space(46)
    If IsNull(.Recordset!CMC7) = True Then
      MsgBox "Existe cheque sem CMC7 cadastrado!"
      Close nArq
      Exit Function
    End If
    
    

    Mid(StrTemp, 1) = Mid(.Recordset!CMC7, 2, 8)
    Mid(StrTemp, 9) = Mid(.Recordset!CMC7, 11, 10)
    Mid(StrTemp, 19) = Mid(.Recordset!CMC7, 22, 12)
    Mid(StrTemp, 31, 16) = Format(.Recordset!Valor, "0000000000000.00")
    
    Print #nArq, StrTemp
    .Recordset.MoveNext
  Loop
  Close #nArq
End With
DepositoBradesco = True
End Function

Private Sub CabecaTodos(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14
StrTemp = "Relção de Cheques Pendentes"
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
StrTemp = "Relção de Cheques para Depósito"
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
Printer.CurrentX = 35 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Banco"
Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Agência"
Printer.CurrentX = 75 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cheque Nr."
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cod.Cli."
Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
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
Dim Quebra As String

With dbCheques
  '.RecordSource = "select *from cheques where compensado=0 and cobrando=0 and codigosoma='1' and protesto=0 order by datacheque, comp, agencia, banco, conta,chequenr, valor"
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
  Quebra = grdCheques.Columns(ColunaQuebra).Text
  Do While .Recordset.EOF = False
    If Quebra <> grdCheques.Columns(ColunaQuebra).Text Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      StrTemp = "Sub-Total: " & Format(SubTotal, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      SubCabeca Largura
      SubTotal = 0
      Quebra = grdCheques.Columns(ColunaQuebra).Text
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
    Printer.CurrentX = 35 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 75 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoCliente
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
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
  .RecordSource = "select *from cheques where compensado=0 and codigosoma='" & codigoSoma & "' order by datacheque, comp, agencia, banco, conta,chequenr, valor"
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
    Printer.CurrentX = 35 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Banco
    Printer.CurrentX = 55 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Agencia
    Printer.CurrentX = 75 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 100 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!chequenr
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoCliente
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
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
    .RecordSource = "select *from cheques where compensado=0 and codigosoma='" & codigoSoma & "'"
    .Refresh
    If .Recordset.RecordCount = 0 Then Exit Sub
    Resposta = MsgBox("Deseja remover os cheques da soma?", vbYesNo)
    If Resposta = vbNo Then Exit Sub
    Do While .Recordset.RecordCount <> 0
      .Recordset.Edit
      .Recordset!codigoSoma = "1"
      .Recordset.Update
      .Refresh
    Loop
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
      Else
        With dbDepositar
          .Refresh
          If StrTemp <> "" Then
            .Recordset.FindFirst StrTemp
            If .Recordset.NoMatch = False Then
              lblData.Caption = Format(.Recordset!datacheque, "short date")
              lblValor.Caption = Format(.Recordset!Valor, "Currency")
              MsgBox "Cheque já somado!"
              Exit Sub
            Else
              MsgBox "Cheque não localizado!"
            End If
          End If
        End With
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
Dim Resposta As Integer

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With


If DateDiff("d", Date, txtData.Value) >= 90 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode depositar com data futura acima de 90 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -10 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode depositar com data anterior a 10 dias!"
    Exit Sub
  End If
End If

If lblTotal.Caption = "" Then Exit Sub
If IsNumeric(lblTotal.Caption) = False Then Exit Sub
If CCur(lblTotal.Caption) = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Sub
End If
If cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Conta inválida!"
  cboConta.SetFocus
  Exit Sub
End If
Resposta = MsgBox("Deseja confirmar o depósito?", vbYesNo)
If Resposta = vbNo Then Exit Sub

codigoSoma = GeraCodigo()

With dbDepositar
  .Refresh
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Selecione pelo menos um cheque!"
    Exit Sub
  End If
  Do While .Recordset.EOF = False
    If IsNull(.Recordset!CodigoCliente) = False Then
      dbClientes.Recordset.FindFirst "codigochequecliente=" & .Recordset!CodigoCliente
    Else
      If IsNull(.Recordset!CPF) = False Then
        dbClientes.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
      Else
        If IsNull(dbDepositar.Recordset!CNPJ) = False Then
          dbClientes.Recordset.FindFirst "cnpj='" & .Recordset!CNPJ & "'"
        Else
          dbClientes.Recordset.FindFirst "codigochequecliente=0"
        End If
      End If
    End If
    If dbClientes.Recordset.NoMatch = False Then
      dbClientes.Recordset.Edit
      dbClientes.Recordset!Depositados = dbClientes.Recordset!Depositados + 1
      dbClientes.Recordset!valordepositado = dbClientes.Recordset!valordepositado + .Recordset!Valor
      If IsNull(dbClientes.Recordset!saldopendente) = True Then dbClientes.Recordset!saldopendente = 0
      dbClientes.Recordset!saldopendente = dbClientes.Recordset!saldopendente - .Recordset!Valor
      If dbClientes.Recordset!saldopendente < 0 Then dbClientes.Recordset!saldopendente = 0
      dbClientes.Recordset.Update
    End If
    
    .Recordset.Edit
    .Recordset!codigoSoma = codigoSoma
    .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
    .Recordset!contadescri = dbContas.Recordset!Descri
    .Recordset!Datacomp = txtData.Value
    .Recordset!compensado = True
    .Recordset.Update
    .Recordset.MoveNext
  Loop
End With

With dbConcilia
  .Recordset.AddNew
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!DataLanc = Now
  .Recordset!Tipo = "Depósito"
  .Recordset!Codigo = 999999995
  .Recordset!Descri = "Depósito de cheques!"
  .Recordset!NrDocumento = codigoSoma
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset.Update
End With
With dbContas
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo + CCur(lblTotal.Caption)
  .Recordset.Update
End With


codigoSoma = "Depositando"

With dbCheques
  .Refresh
  .Refresh
End With
With dbDepositar
  .RecordSource = "select *from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
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
Unload Me

End Sub

Private Sub cmdCustódia_Click()
Dim Resposta As Integer, Custodia As Integer, CodigoPrevisao As Double
Dim StrTemp As String, StrTemp2 As String

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With
With dbDepositar
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    .Recordset.FindFirst "datacheque<=#" & DataInglesa(DateAdd("d", 3, Date)) & "#"
    If .Recordset.NoMatch = False Then
      MsgBox "Existe cheque com data inferior ao prazo limite de 3 dias para efetuar a compensação no dia correto!"
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
If cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Conta inválida!"
  cboConta.SetFocus
  Exit Sub
End If
Resposta = MsgBox("Deseja confirmar a custódia?", vbYesNo)
If Resposta = vbNo Then Exit Sub

'**********************************************************************************************
'                      GERA O ARQUIVO DE CUSTÓDIA PARA O BANCO
'**********************************************************************************************
Custodia = MsgBox("Deseja gerar arquivo para custódia?", vbYesNo)
If Custodia = vbYes Then
  Load frmConciliaDepositoSelecionaBanco
  With frmConciliaDepositoSelecionaBanco
    .lstBanco.Clear
    .lstBanco.AddItem "Banco do Brasil"
    .lstBanco.AddItem "Bradesco"
    .Show vbModal
    Select Case .Banco
      Case "Banco do Brasil"
        If CustodiaBB = False Then
          Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
          Do While Custodia = vbYes
            If CustodiaBB = True Then Exit Do
            Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
          Loop
        End If
      Case "Bradesco"
        If CustodiaBradesco = False Then
          Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
          Do While Custodia = vbYes
            If CustodiaBradesco = True Then
              dbContas.Recordset.Edit
              dbContas.Recordset!custodialote = Lote + 1
              dbContas.Recordset.Update
              Exit Do
            End If
            Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
          Loop
        Else
          dbContas.Recordset.Edit
          dbContas.Recordset!custodialote = Lote + 1
          dbContas.Recordset.Update
        End If
    End Select
  End With
End If
'**********************************************************************************************
Resposta = MsgBox("O arquivo de custódia foi gerado corretamente?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then
  dbContas.Recordset.Edit
  dbContas.Recordset!custodialote = dbContas.Recordset!custodialote - 1
  dbContas.Recordset.Update
  Exit Sub
End If
codigoSoma = GeraCodigo()

With dbDepositar
  StrTemp = .Recordset.Sort
  .Refresh
  .Recordset.Sort = StrTemp
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
      dbClientes.Recordset!Depositados = dbClientes.Recordset!Depositados + 1
      dbClientes.Recordset!valordepositado = dbClientes.Recordset!valordepositado + .Recordset!Valor
      dbClientes.Recordset.Update
    End If
    
    With dbPendencias
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "codigoconta=" & dbContas.Recordset!CodigoConta & " and descri='Custódia de cheques!' and conciliado=0 and data=#" & DataInglesa(Trim(Str(dbDepositar.Recordset!datacheque))) & "#"
        If .Recordset.NoMatch = True Then
          .Recordset.AddNew
          .Recordset!Valor = 0
        Else
          .Recordset.Edit
          If IsNull(.Recordset!NrDoc) = False Then
            codigoSoma = .Recordset!NrDoc
          Else
            codigoSoma = GeraCodigo
          End If
        End If
      Else
        .Recordset.AddNew
        .Recordset!Valor = 0
      End If
      .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
      .Recordset!Conta = dbContas.Recordset!Descri
      .Recordset!CodigoDespesa = 0
      .Recordset!Descri = "Custódia de cheques!"
      .Recordset!NrDoc = codigoSoma
      .Recordset!Valor = .Recordset.Valor + dbDepositar.Recordset!Valor
      .Recordset!Data = dbDepositar.Recordset!datacheque
      CodigoPrevisao = .Recordset!codigopendencia
      .Recordset.Update
    End With
    
    .Recordset.Edit
    .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
    .Recordset!contadescri = dbContas.Recordset!Descri
    .Recordset!Datacomp = .Recordset!datacheque
    .Recordset!codigoSoma = codigoSoma
    .Recordset!codigoprevisaorecebe = CodigoPrevisao
    .Recordset!Custodia = True
    .Recordset.Update
    .Recordset.MoveNext
  Loop
  
End With

StrTemp = codigoSoma
codigoSoma = GeraCodigo()
Do While StrTemp = codigoSoma
  codigoSoma = GeraCodigo()
Loop
With dbCheques
  .Refresh
  .Refresh
End With
With dbDepositar
  .RecordSource = "select *from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
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

Unload Me

End Sub

Private Sub cmdGeraArquivo_Click()
Dim StrTemp As String

If dbDepositar.Recordset.RecordCount = 0 Then
  MsgBox "Selecione pelo menos um cheque!"
  Exit Sub
End If
With dbContas
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboConta.Text = "" Then
    MsgBox "Selecione uma conta!"
    Exit Sub
  End If
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    MsgBox "Selecione uma conta!"
    Exit Sub
  End If
  Call CboConta_LostFocus
  If cboConta.Text <> .Recordset!Descri Then
    MsgBox "Selecione uma conta!"
    Exit Sub
  End If
End With

StrTemp = InputBox("Digite 'C' para Custódia ou 'D' para depósito!")

Select Case UCase(StrTemp)
  Case "C"
    Load frmConciliaDepositoSelecionaBanco
    With frmConciliaDepositoSelecionaBanco
      .lstBanco.Clear
      .lstBanco.AddItem "Banco do Brasil"
      .lstBanco.AddItem "Bradesco"
      .Show vbModal
      Select Case .Banco
        Case "Banco do Brasil"
          If CustodiaBB = False Then
            Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
            Do While Custodia = vbYes
              If CustodiaBB = True Then Exit Do
              Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
            Loop
          End If
        Case "Bradesco"
          If CustodiaBradesco = False Then
            Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
            Do While Custodia = vbYes
              If CustodiaBradesco = True Then
                dbContas.Recordset.Edit
                dbContas.Recordset!custodialote = Lote + 1
                dbContas.Recordset.Update
                Exit Do
              End If
              Custodia = MsgBox("Erro ao gerar o arquivo de custodia! Deseja tentar de novo?", vbYesNoCancel)
            Loop
          Else
            dbContas.Recordset.Edit
            dbContas.Recordset!custodialote = Lote + 1
            dbContas.Recordset.Update
          End If
      End Select
    End With
  Case "D"
    DepositoBradesco
  Case Else
    MsgBox "Opção inválida!"
End Select
End Sub

Private Sub cmdGravar_Click()
Dim Resposta As Integer
Dim StrTemp As String

With dbCheques
  StrTemp = .Recordset.Sort
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione um cheque primeiro!"
    Exit Sub
  End If
  If txtDocumemto.Text = "" Then
    MsgBox "Informe um número de documento!"
    txtDocumemto.SetFocus
    Exit Sub
  End If
  Resposta = MsgBox("Deseja gravar a alteração?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Filter = "comp=" & .Recordset!COMP & " and banco=" & .Recordset!Banco & " and conta='" & .Recordset!Conta & "' and agencia=" & .Recordset!Agencia
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset!CPF = txtDocumemto.Text
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  txtDocumemto.Text = ""
  .Refresh
  .Recordset.Sort = StrTemp
  ProcuraCheque
End With
End Sub

Private Sub cmdImprime_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja imprimir relação de cheques?" & Chr(vbKeyReturn) & "Sim - Imprime todos os cheques," & Chr(vbKeyReturn) & "Não - Imprime soma," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
Select Case Resposta
  Case vbYes
    ImprimeTodosCheques
  Case vbNo
    ImprimeChequesSomados
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
  A = .Recordset!codigocheque
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then Exit Sub
  .Recordset.Edit
  If txtCMC7.Text <> "" Then
    If IsNull(.Recordset!CMC7) = True Then
      .Recordset!CMC7 = txtCMC7.Text
    Else
      If Trim(.Recordset!CMC7) = "" Then
        .Recordset!CMC7 = txtCMC7.Text
      End If
    End If
  End If
  .Recordset!codigoSoma = codigoSoma
  .Recordset.Update
  .Refresh
End With
With dbDepositar
  .Refresh
  If .Recordset.EOF = False Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
qTotalPendente.Refresh
lblCheques.Caption = "Cheques para Depósito: " & dbDepositar.Recordset.RecordCount
lblCheques.Refresh
MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
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
  .Recordset!codigoSoma = "1"
  .Recordset.Update
  .Refresh
  If .Recordset.EOF = False Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With dbCheques
  .Refresh
End With
With qSomaCheques
  .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
lblCheques.Caption = "Cheques para Depósito: " & dbDepositar.Recordset.RecordCount
lblCheques.Refresh
MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

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

txtData.Value = Date

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

codigoSoma = "Depositando"

With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where compensado=0 and custodia=0 and cobrando=0 and codigosoma='1' and protesto=0 order by codigocliente, datacheque, comp, agencia, banco, conta,chequenr, valor"
  .Refresh
End With
With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  ColunaQuebra = 7
  .RecordSource = "select *from cheques where codigosoma='" & codigoSoma & "'"
  .Refresh
End With
With dbPendencias
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
  .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & codigoSoma & "'"
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
  .RecordSource = "select sum(valor) as total from cheques where compensado=0 and custodia=0 and cobrando=0 and codigosoma='1' and protesto=0"
  .Refresh
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

Select Case Usuarios.Grupo.ChequeDeposito
  Case 1 'Somente leitura
    cmdGravar.Enabled = False
    cmdSomar.Enabled = False
    cmdSubtrair.Enabled = False
    cmdConfirma.Enabled = False
    cmdCustódia.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
LimpaSoma
End Sub

Private Sub grdCheques_HeadClick(ByVal ColIndex As Integer)
ColunaQuebra = ColIndex
If dbCheques.RecordSource = "Select *from cheques where codigosoma='1' and compensado=0 and protesto=0 and cobrando=0 order by " & grdCheques.Columns(ColIndex).DataField Then
  dbCheques.RecordSource = "Select *from cheques where codigosoma='1' and compensado=0 and protesto=0 and cobrando=0 order by " & grdCheques.Columns(ColIndex).DataField & " desc"
Else
  dbCheques.RecordSource = "Select *from cheques where codigosoma='1' and compensado=0 and protesto=0 and cobrando=0 order by " & grdCheques.Columns(ColIndex).DataField
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
