VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaProtesto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colocar Cheques em Protesto"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   Icon            =   "frmConciliaProtesto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9465
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   960
      Top             =   2400
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
      Left            =   2400
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.Data dbClientes 
      Caption         =   "dbClientes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from chequesclientes"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   5880
      Width           =   1095
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
      Left            =   6960
      TabIndex        =   6
      Top             =   960
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
      TabIndex        =   5
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   6480
      Top             =   2400
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   8640
      Picture         =   "frmConciliaProtesto.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0 and codigosoma='1'"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbDepositar 
      Caption         =   "dbDepositar"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from cheques where compensado=0"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from previsaorecebimentos"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbDespesa 
      Caption         =   "dbDespesas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from despesaslanc2"
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qSomaCheques 
      Caption         =   "qSomaCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data qTotalPendente 
      Caption         =   "qTotalPendente"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where codigosoma='1' and compensado=0"
      Top             =   5280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid grdCheques 
      Bindings        =   "frmConciliaProtesto.frx":0EC4
      Height          =   1815
      Left            =   120
      OleObjectBlob   =   "frmConciliaProtesto.frx":0EDC
      TabIndex        =   3
      Top             =   1440
      Width           =   9255
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5880
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   8
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
      TabIndex        =   9
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
      TabIndex        =   10
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
      TabIndex        =   11
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
      TabIndex        =   12
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
      Bindings        =   "frmConciliaProtesto.frx":27CE
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmConciliaProtesto.frx":27E8
      TabIndex        =   17
      Top             =   3840
      Width           =   9255
   End
   Begin MSComCtl2.DTPicker txtVencimento 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39647
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Protestado/Resgatado em:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.Label Label8 
      Caption         =   "Leitor de Código de barras:"
      Height          =   255
      Left            =   2400
      TabIndex        =   29
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblCheques 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   5520
      TabIndex        =   28
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label52 
      AutoSize        =   -1  'True
      Caption         =   "Comp:"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   720
      Width           =   450
   End
   Begin VB.Label Label51 
      AutoSize        =   -1  'True
      Caption         =   "Banco:"
      Height          =   195
      Left            =   720
      TabIndex        =   26
      Top             =   720
      Width           =   510
   End
   Begin VB.Label Label50 
      AutoSize        =   -1  'True
      Caption         =   "Agência:"
      Height          =   195
      Left            =   1320
      TabIndex        =   25
      Top             =   720
      Width           =   630
   End
   Begin VB.Label Label49 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   2040
      TabIndex        =   24
      Top             =   720
      Width           =   465
   End
   Begin VB.Label Label48 
      AutoSize        =   -1  'True
      Caption         =   "Cheque:"
      Height          =   195
      Left            =   3000
      TabIndex        =   23
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   720
      Width           =   390
   End
   Begin VB.Label lblData 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   4920
      TabIndex        =   15
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total em Protesto:"
      Height          =   195
      Left            =   3840
      TabIndex        =   22
      Top             =   5880
      Width           =   1530
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6480
      TabIndex        =   21
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total em Cobrança:"
      Height          =   255
      Left            =   6120
      TabIndex        =   20
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label lblTotalPendente 
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
      Height          =   315
      Left            =   7800
      TabIndex        =   19
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   120
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label70 
      Caption         =   "Leitura Automática"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   5880
      Width           =   1815
   End
End
Attribute VB_Name = "frmConciliaProtesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Porta As Integer, codigoSoma As String, strOrdem As String, StrOrdem2 As String

Private Sub ImprimeDados(ByVal Documento As String)
With dbClientes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "cic='" & Documento & "'"
    If .Recordset.NoMatch = True Then
      .Recordset.FindFirst "cnpj='" & Documento & "'"
    End If
    If .Recordset.NoMatch = True Then Exit Sub
  End If
  Y1 = Printer.CurrentY + 1
  
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
StrTemp = "Relção de Cheques em Cobrança"
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
StrTemp = "Relação de Cheques em Protesto"
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
Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cod.Cliente"
Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
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
    If Printer.CurrentY > Printer.ScaleHeight - 20 Then
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
  
  Largura = 190
  Dia = Now
  
  
  CabecaSoma Largura, Dia
  
  Printer.FontSize = 10
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 20 Then
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
    Printer.CurrentX = 145 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!CodigoCliente) = False Then
      StrTemp = .Recordset!CodigoCliente
    Else
      dbClientes.Recordset.FindFirst "cic='" & .Recordset!CPF & "'"
      If dbClientes.Recordset.NoMatch = True Then
        dbClientes.Recordset.FindFirst "cnpj='" & .Recordset!CPF & "'"
        If dbClientes.Recordset.NoMatch = False Then
          .Recordset.Edit
          .Recordset!CodigoCliente = dbClientes.Recordset!codigochequecliente
          .Recordset.Update
          StrTemp = .Recordset!CodigoCliente
        Else
          StrTemp = ""
        End If
      Else
        .Recordset.Edit
        .Recordset!CodigoCliente = dbClientes.Recordset!codigochequecliente
        .Recordset.Update
        StrTemp = .Recordset!CodigoCliente
      End If
    End If
    Printer.CurrentX = 165 - Printer.TextWidth(StrTemp)
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
  
  
  Printer.EndDoc
End With
NaoImprime:


End Sub

Private Sub ImprimeChequesSomados(ByVal SoAtual As Boolean)
Dim Largura As Double, Dia As Date, StrTemp As String
Dim DiaAtual As Date, SubTotal As Currency, Total As Currency

With dbDepositar
  If SoAtual = False Then
    .Refresh
  End If
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
    
    ImprimeDados .Recordset!CPF
    
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
    
    If SoAtual = True Then Exit Do
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

Private Sub cmdImprime_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja imprimir relação de cheques em protesto?" & Chr(vbKeyReturn) & "Sim - Imprime os cheques em protesto," & Chr(vbKeyReturn) & "Não - Imprime cheques devolvidos em cobrança," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
Select Case Resposta
  Case vbYes
    Resposta = MsgBox("Deseja imprimir relatório detalhado de cheques em protesto?" & Chr(vbKeyReturn) & "Sim - Imprime Relatório Detalhado," & Chr(vbKeyReturn) & "Não - Imprime listagem simples," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
    Select Case Resposta
      Case vbYes
        Resposta = MsgBox("Deseja imprimir somente o cheque selecionado?" & Chr(vbKeyReturn) & "Sim - Imprime somente selecionado," & Chr(vbKeyReturn) & "Não - Imprime todos," & Chr(vbKeyReturn) & "Cancela - cancela a operação", vbYesNoCancel)
        If Resposta = vbNo Then
          ImprimeChequesSomados False
        ElseIf Resposta = vbYes Then
          ImprimeChequesSomados True
        Else
          Exit Sub
        End If
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
Dim A As Double, Valor As Currency, StrTemp As String, DataDocumento As Date
Dim CodigoCliente As Double, Nome As String, Documento As String

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

With dbCheques
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  For i = 0 To MaskEdBox1.Count - 1
    If Trim(MaskEdBox1(i).Text) = "" Then
      Resposta = MsgBox("Deseja incluir o cheque atual em protesto?", vbYesNo + vbDefaultButton2)
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
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de cheques! Cheque não incluido!"
    Exit Sub
  End If
  .Recordset.Edit
  DataDocumento = .Recordset!datacheque
  CodigoCliente = .Recordset!CodigoCliente
  .Recordset!protesto = True
  .Recordset!dataprotesto = Now
  StrTemp = .Recordset!COMP & "/" & .Recordset!Banco & "/" & .Recordset!Agencia & "/" & .Recordset!Conta & "/" & .Recordset!chequenr & " - Cod:" & .Recordset!CodigoCliente
  Valor = -.Recordset!Valor
  .Recordset.Update
  .Refresh
End With
With dbClientes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codigochequecliente=" & CodigoCliente
    If .Recordset.NoMatch = False Then
      Nome = .Recordset!Nome
      If IsNull(.Recordset!CIC) = False Then
        If .Recordset!CIC <> "" Then
          Documento = .Recordset!CIC
        End If
      End If
      If IsNull(.Recordset!CNPJ) = False Then
        If .Recordset!CNPJ <> "" Then
          Documento = .Recordset!CNPJ
        End If
      End If
    End If
  End If
End With
With dbDespesa
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Despesa"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoDespesa = 0
  .Recordset!NrDocumento = A
  .Recordset!Descri = "Protesto de Cheque"
  .Recordset!Obs = StrTemp
  .Recordset!Valor = Valor
  .Recordset!valorpago = Valor
  .Recordset!compensado = True
  .Recordset!fechamentodiario = True
  .Recordset!codigoenviar = "1"
  .Recordset.Update
End With

Resposta = vbYes
Do While Resposta = vbYes
  If ProtestoAdicionaHistorico(txtVencimento.Value, DataDocumento, "Cheque", "Protestando", StrTemp, Valor, Nome, Documento, , CodigoCliente) = False Then
    Resposta = MsgBox("Não foi possível criar o histórico de protesto. Deseja Tentar de novo?", vbYesNo)
  Else
    Exit Do
  End If
Loop

With dbDepositar
  .Refresh
  If .Recordset.EOF = False Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotalPendente.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=-1 and EmpresaDeCobranca=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotal.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
lblCheques.Caption = dbDepositar.Recordset.RecordCount

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "
txtCMC7.Text = ""

txtCMC7.SetFocus
End Sub

Private Sub cmdSubtrair_Click()
Dim A As Double, Resposta As Integer, DataDocumento As Date, Valor As Currency
Dim CodigoCliente As Double, Nome As String, Documento As String

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With



With dbDepositar
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  Resposta = MsgBox("Deseja remover o cheque atual do protesto?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  
  A = .Recordset!codigocheque
  CodigoCliente = .Recordset!CodigoCliente
  With dbDespesa
    .RecordSource = "select *from despesaslanc2"
    .Refresh
    .Recordset.FindFirst "nrdocumento='" & A & "' and Descri='Protesto de Cheque'"
    If .Recordset.NoMatch = True Then
      .RecordSource = "select *from despesaslanc2"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "nrdocumento='" & A & "' and Descri='Protesto de Cheque'"
        If .Recordset.NoMatch = True Then
          .RecordSource = "select *from despesaslanc2"
          .Refresh
          dbDepositar.Refresh
          dbDepositar.Recordset.FindFirst "codigocheque=" & A
          If dbDepositar.Recordset.NoMatch = True Then
            MsgBox "Erro na tabela de cheques!"
            Exit Sub
          End If
          .Recordset.AddNew
          .Recordset!CodigoFechamento = 0
          .Recordset!Origem = "Despesa"
          .Recordset!Data = Date
          .Recordset!Hora = Now
          .Recordset!Vencimento = Date
          .Recordset!CodigoDespesa = 0
          .Recordset!NrDocumento = A
          .Recordset!Descri = "Protesto de Cheque Resgatado"
          StrTemp = dbDepositar.Recordset!COMP & "/" & dbDepositar.Recordset!Banco & "/" & dbDepositar.Recordset!Agencia & "/" & dbDepositar.Recordset!Conta & "/" & dbDepositar.Recordset!chequenr
          .Recordset!Obs = StrTemp
          .Recordset!Valor = dbDepositar.Recordset!Valor
          .Recordset!valorpago = dbDepositar.Recordset!Valor
          .Recordset!compensado = True
          .Recordset!fechamentodiario = True
          On Error Resume Next
          .Recordset!codigoenviar = "1"
          On Error GoTo 0
          .Recordset.Update
        Else
          If .Recordset!Fechamento = True Then
            .RecordSource = "select *from despesaslanc2"
            .Refresh
            dbDepositar.Refresh
            dbDepositar.Recordset.FindFirst "codigocheque=" & A
            If dbDepositar.Recordset.NoMatch = True Then
              MsgBox "Erro na tabela de cheques!"
              Exit Sub
            End If
            .Recordset.AddNew
            .Recordset!CodigoFechamento = 0
            .Recordset!Origem = "Despesa"
            .Recordset!Data = Date
            .Recordset!Hora = Now
            .Recordset!Vencimento = Date
            .Recordset!CodigoDespesa = 0
            .Recordset!NrDocumento = A
            .Recordset!Descri = "Protesto de Cheque Resgatado"
            StrTemp = dbDepositar.Recordset!COMP & "/" & dbDepositar.Recordset!Banco & "/" & dbDepositar.Recordset!Agencia & "/" & dbDepositar.Recordset!Conta & "/" & dbDepositar.Recordset!chequenr
            .Recordset!Obs = StrTemp
            .Recordset!Valor = dbDepositar.Recordset!Valor
            .Recordset!valorpago = dbDepositar.Recordset!Valor
            .Recordset!compensado = True
            .Recordset!fechamentodiario = True
            .Recordset!codigoenviar = "1"
            .Recordset.Update
          Else
            .Recordset.Delete
          End If
        End If
      Else
        MsgBox "Erro na tabela de despesas!"
        Exit Sub
      End If
    Else
      If .Recordset!Fechamento = True Then
        .RecordSource = "select *from despesaslanc2"
        .Refresh
        dbDepositar.Refresh
        dbDepositar.Recordset.FindFirst "codigocheque=" & A
        If dbDepositar.Recordset.NoMatch = True Then
          MsgBox "Erro na tabela de cheques!"
          Exit Sub
        End If
        .Recordset.AddNew
        .Recordset!CodigoFechamento = 0
        .Recordset!Origem = "Despesa"
        .Recordset!Data = Date
        .Recordset!Hora = Now
        .Recordset!Vencimento = Date
        .Recordset!CodigoDespesa = 0
        .Recordset!NrDocumento = A
        .Recordset!Descri = "Protesto de Cheque Resgatado"
        StrTemp = dbDepositar.Recordset!COMP & "/" & dbDepositar.Recordset!Banco & "/" & dbDepositar.Recordset!Agencia & "/" & dbDepositar.Recordset!Conta & "/" & dbDepositar.Recordset!chequenr
        .Recordset!Obs = StrTemp
        .Recordset!Valor = dbDepositar.Recordset!Valor
        .Recordset!valorpago = dbDepositar.Recordset!Valor
        .Recordset!compensado = True
        .Recordset!fechamentodiario = True
        On Error Resume Next
        .Recordset!codigoenviar = "1"
        On Error GoTo 0
        .Recordset.Update
      Else
        .Recordset.Delete
      End If
    End If
  End With
  If .Recordset.EOF = False Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
  .Refresh
  .Recordset.FindFirst "codigocheque=" & A
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de cheques em protesto!"
    Exit Sub
  End If
  .Recordset.Edit
  DataDocumento = .Recordset!datacheque
  StrTemp = .Recordset!COMP & "/" & .Recordset!Banco & "/" & .Recordset!Agencia & "/" & .Recordset!Conta & "/" & .Recordset!chequenr & " - Cod:" & .Recordset!CodigoCliente
  Valor = .Recordset!Valor
  .Recordset!protesto = False
  .Recordset.Update
  .Refresh
End With

With dbClientes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codigochequecliente=" & CodigoCliente
    If .Recordset.NoMatch = False Then
      Nome = .Recordset!Nome
      If IsNull(.Recordset!CIC) = False Then
        If .Recordset!CIC <> "" Then
          Documento = .Recordset!CIC
        End If
      End If
      If IsNull(.Recordset!CNPJ) = False Then
        If .Recordset!CNPJ <> "" Then
          Documento = .Recordset!CNPJ
        End If
      End If
    End If
  End If
End With

Resposta = vbYes
Do While Resposta = vbYes
  If ProtestoAdicionaHistorico(txtVencimento.Value, DataDocumento, "Cheque", "Resgatando", StrTemp, Valor, Nome, Documento, , CodigoCliente) = False Then
    Resposta = MsgBox("Não foi possível criar o histórico de protesto. Deseja Tentar de novo?", vbYesNo)
  Else
    Exit Do
  End If
Loop

With dbCheques
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotalPendente.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=-1 and EmpresaDeCobranca=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotal.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
lblCheques.Caption = dbDepositar.Recordset.RecordCount

MaskEdBox1(0).Text = "   "
MaskEdBox1(1).Text = "   "
MaskEdBox1(2).Text = "    "
MaskEdBox1(3).Text = "      - "
MaskEdBox1(4).Text = "      "

MaskEdBox1(0).SetFocus

End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField & ", chequenr, valor" Then
  strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField & " desc, chequenr, valor"
Else
  strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField & ", chequenr, valor"
End If

With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=-1 and EmpresaDeCobranca=0" & strOrdem
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

txtVencimento.Value = Date


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
strOrdem = " order by CodigoCliente, chequenr, valor"
StrOrdem2 = " order by CodigoCliente, chequenr, valor"
With dbCheques
  On Error Resume Next
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques order by EmpresaDeCobranca"
  .Refresh
  If Err.Number <> 0 Then
    On Error GoTo 0
    On Error Resume Next
    Set Ws = DBEngine.Workspaces(0)
    Set db = Ws.OpenDatabase(Caminho)
    db.Execute "alter table Cheques Add column EmpresaDeCobranca bit;"
    If Err.Number <> 0 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " atualizando a tabela de Cheques"
      End
    End If
    On Error GoTo 0
    On Error Resume Next
    db.Execute "alter table Cheques2 Add column EmpresaDeCobranca bit;"
    If Err.Number <> 0 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " atualizando a tabela de Cheques2"
      End
    End If
    On Error GoTo 0
    On Error Resume Next
    db.Execute "alter table Cheques Add column DataEmpresaDeCobranca datetime;"
    If Err.Number <> 0 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " atualizando a tabela de Cheques"
      End
    End If
    On Error GoTo 0
    On Error Resume Next
    db.Execute "alter table Cheques2 Add column DataEmpresaDeCobranca datetime;"
    If Err.Number <> 0 Then
      MsgBox "Erro " & Err.Number & " - " & Err.Description & " atualizando a tabela de Cheques2"
      End
    End If
  End If
End With



With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=0" & StrOrdem2
  .Refresh
End With

With dbDepositar
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=-1 and EmpresaDeCobranca=0" & strOrdem
  .Refresh
End With
With dbPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesa
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotalPendente.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
With qTotalPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from cheques where cobrando=-1 and protesto=-1 and EmpresaDeCobranca=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    If IsNumeric(.Recordset!Total) = True Then
      lblTotal.Caption = Format(.Recordset!Total, "currency")
    End If
  End If
End With
With dbClientes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

lblCheques.Caption = dbDepositar.Recordset.RecordCount

Select Case Usuarios.Grupo.ControleCartoes
  Case 1 'Somente leitura
    cmdSomar.Enabled = False
    cmdSubtrair.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub grdCheques_DblClick()
With grdCheques
  .AllowUpdate = True
  If .Col <> 5 Then
    .Col = 5
  End If
End With
End Sub

Private Sub grdCheques_HeadClick(ByVal ColIndex As Integer)
If StrOrdem2 = " order by " & grdCheques.Columns(ColIndex).DataField & ", chequenr, valor" Then
  StrOrdem2 = " order by " & grdCheques.Columns(ColIndex).DataField & " desc, chequenr, valor"
Else
  StrOrdem2 = " order by " & grdCheques.Columns(ColIndex).DataField & ", chequenr, valor"
End If

With dbCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from cheques where cobrando=-1 and protesto=0" & StrOrdem2
  .Refresh
End With
End Sub

Private Sub grdCheques_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdCheques.Col <> 5 Then
  grdCheques.Col = 5
End If
End Sub

Private Sub GrdDeposito_HeadClick(ByVal ColIndex As Integer)
dbDepositar.Refresh
dbDepositar.Recordset.Sort = GrdDeposito.Columns(ColIndex).DataField
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
    txtVencimento.SetFocus
  End With
End If

End Sub

Private Sub txtVencimento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtVencimento_LostFocus()
Me.KeyPreview = True
End Sub
