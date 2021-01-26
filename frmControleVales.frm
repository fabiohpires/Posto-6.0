VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmControleVales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Vales e Comissões de Funcionários"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "frmControleVales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4815
      Left            =   5640
      TabIndex        =   4
      Top             =   5520
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Data dbConciliaNova 
         Caption         =   "dbConciliaNova"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from conciliaNova"
         Top             =   720
         Width           =   3180
      End
      Begin VB.Data dbCartoes 
         Caption         =   "dbCartoes"
         Connect         =   "Access 2000;"
         DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Cartoes"
         Top             =   360
         Width           =   3180
      End
      Begin VB.Data dbContas 
         Caption         =   "dbContas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Contas"
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Data dbDespesaLanc 
         Caption         =   "dbDespesaLanc"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "DespesasLanc2"
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Data qValesTotal 
         Caption         =   "qValesTotal"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select QVales.*, fechamentodecaixa.* from qvales, fechamentodecaixa where qvales.codigocaixa=fechamentodecaixa.codigofechamento"
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Data qVendasTotal 
         Caption         =   "qVendasTotal"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QVendas"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Data dbDespesa 
         Caption         =   "dbDespesa"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from DespesaTipo order by descri"
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Data dbFormaDePg 
         Caption         =   "dbFormaDePg"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from FormaDePagamento order by descri"
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Data dbFormaDePgRecebido 
         Caption         =   "dbFormaDePgRecebido"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QVendas"
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Data qVendas 
         Caption         =   "qVendas"
         Connect         =   "Access 2000;"
         DatabaseName    =   "d:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno"
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Data dbVenda 
         Caption         =   "dbVenda"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Venda2"
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Data dbFuncionarios 
         Caption         =   "dbFuncionarios"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from Vendedores order by nome"
         Top             =   360
         Width           =   3255
      End
      Begin VB.Data qVales 
         Caption         =   "qVales"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "QValesCaixa"
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Data dbVales 
         Caption         =   "dbVales"
         Connect         =   "Access 2000;"
         DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Vales"
         Top             =   720
         Width           =   3255
      End
      Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
         Height          =   330
         Left            =   2880
         Top             =   1080
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Vales"
      TabPicture(0)   =   "frmControleVales.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label86"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label28"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblValesTotal"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboDespesa"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDataBordero"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DBGrid1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cboFormaDePg"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdDespesa"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdIncluirRecebimento"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtOperacoes"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdImprime"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdReceberTodos"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtObs"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Comissões"
      TabPicture(1)   =   "frmControleVales.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdImprime2"
      Tab(1).Control(1)=   "cmdPagarComissoes"
      Tab(1).Control(2)=   "DBGrid2"
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(4)=   "lblTotalComissao"
      Tab(1).ControlCount=   5
      Begin VB.TextBox txtObs 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   15
         Top             =   4440
         Width           =   6375
      End
      Begin VB.CommandButton cmdReceberTodos 
         Caption         =   "Receber Todos"
         Height          =   495
         Left            =   7680
         TabIndex        =   19
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdImprime2 
         Height          =   495
         Left            =   -71400
         Picture         =   "frmControleVales.frx":047A
         Style           =   1  'Graphical
         TabIndex        =   31
         Tag             =   "Imprimir"
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton cmdImprime 
         Height          =   495
         Left            =   9000
         Picture         =   "frmControleVales.frx":0EFC
         Style           =   1  'Graphical
         TabIndex        =   30
         Tag             =   "Imprimir"
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton cmdPagarComissoes 
         Caption         =   "Pagar Comissões"
         Height          =   375
         Left            =   -74760
         TabIndex        =   29
         Top             =   4200
         Width           =   1455
      End
      Begin VB.TextBox txtOperacoes 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
         Height          =   285
         Left            =   5640
         TabIndex        =   14
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton cmdIncluirRecebimento 
         Caption         =   "Receber"
         Height          =   495
         Left            =   6600
         TabIndex        =   17
         Top             =   4200
         Width           =   975
      End
      Begin VB.CommandButton cmdDespesa 
         Caption         =   "Despesa"
         Height          =   375
         Left            =   4200
         TabIndex        =   10
         Top             =   3120
         Width           =   975
      End
      Begin MSDBCtls.DBCombo cboFormaDePg 
         Bindings        =   "frmControleVales.frx":197E
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   3840
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         Text            =   ""
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmControleVales.frx":1998
         Height          =   2415
         Left            =   120
         OleObjectBlob   =   "frmControleVales.frx":19AD
         TabIndex        =   16
         Top             =   480
         Width           =   9855
      End
      Begin MSComCtl2.DTPicker txtDataBordero 
         Height          =   315
         Left            =   4200
         TabIndex        =   13
         Top             =   3840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   115539969
         CurrentDate     =   37600
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmControleVales.frx":2A50
         Height          =   3615
         Left            =   -74880
         OleObjectBlob   =   "frmControleVales.frx":2A66
         TabIndex        =   22
         Top             =   480
         Width           =   9855
      End
      Begin MSDBCtls.DBCombo cboDespesa 
         Bindings        =   "frmControleVales.frx":41C1
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "Descri"
         Text            =   ""
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Operações:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   4200
         Width           =   825
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   -68280
         TabIndex        =   28
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblTotalComissao 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -67200
         TabIndex        =   27
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Despesa:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblValesTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6720
         TabIndex        =   24
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Total:"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Forma de Pagamento:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Operações:"
         Height          =   195
         Left            =   5640
         TabIndex        =   20
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label Label86 
         Caption         =   "Data Borderô:"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   3600
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   7920
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.CheckBox chkJaPago 
      Caption         =   "Já pagos"
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo cboFuncionario 
      Bindings        =   "frmControleVales.frx":41D9
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   115539969
      CurrentDate     =   38286
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   115539969
      CurrentDate     =   38286
   End
   Begin VB.Label Label3 
      Caption         =   "Funcionário:"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Período de lançamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frmControleVales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdemVale As String, StrOrdemComissao As String
Private Sub cboDespesa_LostFocus()
With dbDespesa
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboDespesa.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboDespesa.Text & "'"
  If .Recordset.NoMatch = False Then
    cboDespesa.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboFormaDePg_LostFocus()
With dbFormaDePG
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFormaDePg.Text = "" Then Exit Sub
  .Recordset.FindFirst "descri='" & cboFormaDePg.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFormaDePg.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboFuncionario_LostFocus()
With dbFuncionarios
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboFuncionario.Text = "" Then Exit Sub
  .Recordset.FindFirst "nome='" & cboFuncionario.Text & "'"
  If .Recordset.NoMatch = False Then
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub cmdDespesa_Click()
Dim Valor As Currency

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If qVales.Recordset.EOF = True Then
  MsgBox "Selecione um vale para ser lançado na despesa!"
  Exit Sub
End If
If qVales.Recordset!Cobrado = True Then
  MsgBox "Esse vale já foi cobrado!"
  Exit Sub
End If
Valor = qVales.Recordset!Valor

If cboDespesa.Text <> dbDespesa.Recordset!Descri Then
  MsgBox "Selecione uma despesa válida!", vbCritical, "Erro!"
  cboDespesa.SetFocus
  Exit Sub
End If
With dbVales
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Erro na tabela de vales!"
    Exit Sub
  End If
  .Recordset.FindFirst "codigovale=" & qVales.Recordset!codigovale
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de vales!"
    Exit Sub
  End If
  .Recordset.Edit
  .Recordset!Cobrado = True
  .Recordset!cobradoem = Now
  .Recordset.Update
End With
With dbDespesaLanc
  .Recordset.AddNew
  .Recordset("codigofechamento") = 0
  .Recordset!Origem = "Controle de Vale"
  .Recordset("data") = Date
  .Recordset!Vencimento = Date
  .Recordset("hora") = Now
  .Recordset("codigoconta") = -1
  .Recordset("conta") = "Controle de Vale"
  .Recordset("codigodespesa") = dbDespesa.Recordset("codigodespesa")
  .Recordset("descri") = "Vale " & dbDespesa.Recordset("descri")
  .Recordset("obs") = qVales.Recordset!Nome
  .Recordset!compensado = True
  .Recordset("valor") = -Valor
  .Recordset!valorpago = -Valor
  .Recordset!fechamentodiario = True
  .Recordset!codigoenviar = "1"
  .Recordset.Update
  .Refresh
End With
Call cmdExibir_Click
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String, StrTempTotal As String
Dim strComissao As String, StrComissaoTotal As String

StrTemp = "select *from QValesCaixa where qvales.fechado=-1"
StrTempTotal = "select sum(valor) as total from qvalescaixa where qvales.fechado=-1"
strComissao = "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno and fechamentodiario=-1"
StrComissaoTotal = "select sum(valorcomissao) as total from qvendascomissoes where fechamentodiario=-1"

If chkJaPago.Value = vbChecked Then
  StrTemp = StrTemp & " and datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  StrTempTotal = StrTempTotal & " and datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  strComissao = strComissao & " and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  StrComissaoTotal = StrComissaoTotal & " and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
Else
  StrTemp = StrTemp & " and datacaixa <=#" & DataInglesa(txtDataFim.Value) & "#"
  StrTempTotal = StrTempTotal & " and datacaixa <=#" & DataInglesa(txtDataFim.Value) & "#"
  strComissao = strComissao & " and data <=#" & DataInglesa(txtDataFim.Value) & "#"
  StrComissaoTotal = StrComissaoTotal & " and data <=#" & DataInglesa(txtDataFim.Value) & "#"
End If
If cboFuncionario.Text <> "" Then
  If dbFuncionarios.Recordset.EOF = False Then
    If dbFuncionarios.Recordset!Nome = cboFuncionario.Text Then
      StrTemp = StrTemp & " and codfun=" & dbFuncionarios.Recordset!codigovendedor
      StrTempTotal = StrTempTotal & " and codfun=" & dbFuncionarios.Recordset!codigovendedor
      strComissao = strComissao & " and codigopagamento=" & dbFuncionarios.Recordset!codigovendedor
      StrComissaoTotal = StrComissaoTotal & " and codigopagamento=" & dbFuncionarios.Recordset!codigovendedor
    End If
  End If
End If
If chkJaPago.Value = vbChecked Then
  StrTemp = StrTemp & " and cobrado=-1"
  StrTempTotal = StrTempTotal & " and cobrado=-1"
  strComissao = strComissao & " and pago=-1"
  StrComissaoTotal = StrComissaoTotal & " and pago=-1"
Else
  StrTemp = StrTemp & " and cobrado=0"
  StrTempTotal = StrTempTotal & " and cobrado=0"
  strComissao = strComissao & " and pago=0"
  StrComissaoTotal = StrComissaoTotal & " and pago=0"
End If

With qVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTemp & StrOrdemVale
  .Refresh
End With
With qValesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrTempTotal
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblValesTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblValesTotal.Caption = Format(0, "Currency")
  End If
End With

With qVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = strComissao & StrOrdemComissao
  .Refresh
End With
With qVendasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = StrComissaoTotal
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalComissao.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalComissao.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, StrTemp2 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = Format(Now, "long date") & " " & Format(Now, "short time")
StrTemp2 = "Período: " & txtDataIni.Value & " a " & txtDataFim.Value
If cboFuncionario.Text <> "" Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Funcionário: " & cboFuncionario.Text
End If
If chkJaPago.Value = vbChecked Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Valores já cobrados"
End If

ImprimeGrid DBGrid1, Printer, qVales, 4, , , , , , "Vales de Funcionários", StrTemp, StrTemp2

Printer.EndDoc
NaoImprime:
End Sub

Private Sub cmdImprime2_Click()
Dim StrTemp As String, StrTemp2 As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = Format(Now, "long date") & " " & Format(Now, "short time")
StrTemp2 = "Período: " & txtDataIni.Value & " a " & txtDataFim.Value
If cboFuncionario.Text <> "" Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Funcionário: " & cboFuncionario.Text
End If
If chkJaPago.Value = vbChecked Then
  StrTemp2 = StrTemp2 & Chr(vbKeyReturn) & "Valores já cobrados"
End If

ImprimeGrid DBGrid2, Printer, qVendas, 4, , , , 6, 7, "Comissões", StrTemp, StrTemp2

Printer.EndDoc
NaoImprime:
End Sub

Private Sub cmdIncluirRecebimento_Click()
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency, DescontoPorcento As Currency

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtDataBordero.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If DateDiff("d", Date, txtDataBordero.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtDataBordero.Value) <= -10 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber com data anterior a 10 dias!"
    Exit Sub
  End If
End If


If qVales.Recordset.EOF = True Then
  MsgBox "Informe um vale a ser cobrado!"
  Exit Sub
End If
If qVales.Recordset!Cobrado = True Then
  MsgBox "Este vale já foi cobrado antes!"
  Exit Sub
End If
If cboFormaDePg.Text <> dbFormaDePG.Recordset!Descri Then
  MsgBox "Escolha uma forma de Pagamento válida!", vbCritical, "Erro!"
  cboFormaDePg.SetFocus
  Exit Sub
End If

Tarifa = dbFormaDePG.Recordset!descontovalor
Operacao = dbFormaDePG.Recordset!descontoporoperacao
Porcento = dbFormaDePG.Recordset!DescontoPorcento / 100

TotalOper = 0
If Operacao <> 0 Then
  If IsNumeric(txtOperacoes.Text) = True Then
    TotalOper = CDbl(txtOperacoes.Text)
    If TotalOper = 0 Then
      MsgBox "Informe um valor correto para desconto por operação!"
      txtOperacoes.SetFocus
      Exit Sub
    Else
      Operacao = Operacao * TotalOper
    End If
  Else
    MsgBox "Informe um valor correto para desconto por operação!"
    txtOperacoes.SetFocus
    Exit Sub
  End If
End If
ValorBruto = qVales.Recordset!Valor

If Porcento <> 0 Then
  DescontoPorcento = ValorBruto * Porcento
End If

Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao

With dbVales
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "Erro na tabela de Vales"
    Exit Sub
  End If
  .Recordset.FindFirst "codigovale=" & qVales.Recordset!codigovale
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de Vales"
    Exit Sub
  End If
  .Recordset.Edit
  .Recordset!Cobrado = True
  .Recordset!cobradoem = Now
  .Recordset.Update
End With

With dbFormaDePG
  LucroVenda = LucroVenda + ValorBruto
  Dias = .Recordset("reembolso")
  Mes = .Recordset("mes")
  If Mes = True Then
    Intervalo = "m"
  Else
    Intervalo = "d"
  End If
  If Dias > 0 Then
    ReceberData = DateAdd(Intervalo, Dias, txtDataBordero.Value)
  Else
    Dias = .Recordset("diadomes")
    If Dias > 0 Then
      If Dias >= txtDataBordero.Day Then
        If Dias < 28 Then
          StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
        Else
          StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
          Do While IsDate(StrTemp) = False
            Dias = Dias - 1
            If Dias <= 0 Then Dias = 31
            StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
          Loop
        End If
        ReceberData = CDate(StrTemp)
      Else
        If Dias < 28 Then
          StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
        Else
          StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
          Do While IsDate(StrTemp) = False
            Dias = Dias - 1
            If Dias <= 0 Then Dias = 31
            StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
          Loop
        End If
        ReceberData = CDate(StrTemp)
      End If
    End If
  End If
  If Dias > 0 Then
    Select Case Weekday(ReceberData)
      Case 1 'domingo
        ReceberData = DateAdd("d", 1, ReceberData)
      Case 7 'sábado
        ReceberData = DateAdd("d", 2, ReceberData)
    End Select
    dbCartoes.Refresh
    dbCartoes.Recordset.AddNew
    dbCartoes.Recordset!ValorBruto = 0
    dbCartoes.Recordset!valorliquido = 0
    dbCartoes.Recordset!CodigoConta = .Recordset("codigoconta")
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
    dbCartoes.Recordset!Conta = dbContas.Recordset("descri")
    dbCartoes.Recordset!CodigoFormaPg = .Recordset!CodigoPagamento
    dbCartoes.Recordset!Grupo = .Recordset!Grupo
    dbCartoes.Recordset!Descri = .Recordset("descri")
    dbCartoes.Recordset!DataLanc = txtDataBordero.Value
    dbCartoes.Recordset!DataPrevista = ReceberData
    dbCartoes.Recordset!ValorBruto = dbCartoes.Recordset!ValorBruto + ValorBruto
    dbCartoes.Recordset!valorliquido = dbCartoes.Recordset!valorliquido + Liquido
    If cboFuncionario.Text = "" Then
      StrTemp = qVales.Recordset!Nome
    Else
      StrTemp = cboFuncionario.Text
    End If
    dbCartoes.Recordset!Obs = StrTemp & "-" & txtObs.Text
    dbCartoes.Recordset.Update
    
  Else
    ReceberData = txtDataBordero.Value
    Select Case Weekday(ReceberData)
      Case 1 'domingo
        ReceberData = DateAdd("d", 1, ReceberData)
      Case 7 'sábado
        ReceberData = DateAdd("d", 2, ReceberData)
    End Select
    
    With dbConciliaNova
      .Recordset.AddNew
      .Recordset!CodigoConta = dbFormaDePG.Recordset("codigoconta")
      .Recordset!DataLanc = Now
      .Recordset!compensado = True
      .Recordset!Data = Date
      .Recordset!Tipo = "Vale"
      .Recordset!Codigo = 999999998
      .Recordset!Descri = "Vale recebido de " & qVales.Recordset!Nome
      .Recordset!NrDocumento = Format(txtDataBordero.Value, "short date")
      .Recordset!Valor = ValorBruto
      .Recordset.Update
    End With
    dbContas.Refresh
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
    If dbContas.Recordset.NoMatch = True Then
      MsgBox "Conta " & .Recordset("contas.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
    Else
      TempValor = ValorBruto
      dbContas.Recordset.Edit
      dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + TempValor
      dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
      dbContas.Recordset.Update
    End If
  End If
  
End With

With dbFormaDePgRecebido
  .RecordSource = "formadepagamentorecebido2"
  .Refresh
  .Recordset.AddNew
  .Recordset("codigofechamento") = 0
  .Recordset("codigoformadepg") = dbFormaDePG.Recordset("codigoPagamento")
  .Recordset("descri") = dbFormaDePG.Recordset("descri")
  .Recordset("valorbruto") = ValorBruto
  .Recordset("valordescoper") = Operacao
  .Recordset("valordesctarifa") = Tarifa
  .Recordset("valordesconto") = DescontoPorcento
  .Recordset("valor") = Liquido
  .Recordset("operacoes") = TotalOper
  .Recordset("data") = txtDataBordero.Value
  .Recordset("hora") = Now
  .Recordset!fechamentodiario = True
  .Recordset.Update
  .Refresh
End With
Call cmdExibir_Click
End Sub

Private Sub cmdPagarComissoes_Click()
Dim Resposta As Integer, Total As Currency

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= Date And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With


Total = 0
Call cmdExibir_Click
If chkJaPago.Value = vbChecked Then
  MsgBox "Comissões já pagas, não podem ser pagas novamente!"
  Exit Sub
End If
If cboFuncionario.Text = "" Then
  MsgBox "É preciso selecionar um funcionário para pagar!"
  cboFuncionario.SetFocus
  Exit Sub
Else
  If dbFuncionarios.Recordset.EOF = True Then
    MsgBox "É preciso selecionar um funcionário para pagar!"
    cboFuncionario.SetFocus
    Exit Sub
  Else
    If dbFuncionarios.Recordset!Nome <> cboFuncionario.Text Then
      MsgBox "É preciso selecionar um funcionário para pagar!"
      cboFuncionario.SetFocus
      Exit Sub
    End If
  End If
End If

Total = -CCur(lblTotalComissao.Caption)
Resposta = MsgBox("Deseja pagar as comissões exibidas na tela?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With qVendas
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe comissão para ser paga!"
    Exit Sub
  Else
    .Recordset.MoveLast
    .Recordset.MoveFirst
    dbVenda.Refresh
    If dbVenda.Recordset.RecordCount = 0 Then
      MsgBox "Erro na tabela de Venda!"
      Exit Sub
    End If
    Do While .Recordset.EOF = False
      dbVenda.Recordset.FindFirst "codigovenda=" & .Recordset!codigovenda
      If dbVenda.Recordset.NoMatch = False Then
        dbVenda.Recordset.Edit
        dbVenda.Recordset!Pago = True
        dbVenda.Recordset.Update
      End If
      .Recordset.MoveNext
    Loop
  End If
End With

With dbDespesaLanc
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Despesa"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoConta = 0
  .Recordset!CodigoDespesa = 0
  .Recordset!Descri = "Comissões-" & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
  .Recordset!Obs = cboFuncionario.Text
  .Recordset!Valor = Total
  .Recordset!Fechamento = True
  .Recordset!codigoenviar = "1"
  .Recordset.Update
End With
Call cmdExibir_Click

End Sub

Private Sub cmdReceberTodos_Click()
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency, DescontoPorcento As Currency

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtDataBordero.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With


If chkJaPago.Value = vbChecked Then
  MsgBox "Não pode receber vales já pagos!"
  Exit Sub
End If
If cboFuncionario.Text = "" Then
  MsgBox "Só é permitido receber o valor total por funcionário!"
  Exit Sub
End If
If dbFuncionarios.Recordset.EOF = True Then
  MsgBox "Só é permitido receber o valor total por funcionário!"
  Exit Sub
End If
If dbFuncionarios.Recordset!Nome <> cboFuncionario.Text Then
  MsgBox "Escolha um funcionário correto!"
  Exit Sub
End If
Call cmdExibir_Click
If qVales.Recordset.EOF = True Then
  MsgBox "Informe um vale a ser cobrado!"
  Exit Sub
End If
If qVales.Recordset!Cobrado = True Then
  MsgBox "Este vale já foi cobrado antes!"
  Exit Sub
End If
If cboFormaDePg.Text <> dbFormaDePG.Recordset!Descri Then
  MsgBox "Escolha uma forma de Pagamento válida!", vbCritical, "Erro!"
  cboFormaDePg.SetFocus
  Exit Sub
End If

Tarifa = dbFormaDePG.Recordset!descontovalor
Operacao = dbFormaDePG.Recordset!descontoporoperacao
Porcento = dbFormaDePG.Recordset!DescontoPorcento / 100

TotalOper = 0
If Operacao <> 0 Then
  If IsNumeric(txtOperacoes.Text) = True Then
    TotalOper = CDbl(txtOperacoes.Text)
    If TotalOper = 0 Then
      MsgBox "Informe um valor correto para desconto por operação!"
      txtOperacoes.SetFocus
      Exit Sub
    Else
      Operacao = Operacao * TotalOper
    End If
  Else
    MsgBox "Informe um valor correto para desconto por operação!"
    txtOperacoes.SetFocus
    Exit Sub
  End If
End If
If IsNumeric(lblValesTotal.Caption) = False Then
  MsgBox "O total não é um valor válido!"
  Exit Sub
End If
ValorBruto = CCur(lblValesTotal.Caption)
If Porcento <> 0 Then
  DescontoPorcento = ValorBruto * Porcento
End If

Liquido = ValorBruto - DescontoPorcento - Tarifa - Operacao

qVales.Recordset.MoveLast
qVales.Recordset.MoveFirst
Do While qVales.Recordset.EOF = False
  With dbVales
    .Refresh
    If .Recordset.RecordCount = 0 Then
      MsgBox "Erro na tabela de Vales"
      Exit Sub
    End If
    .Recordset.FindFirst "codigovale=" & qVales.Recordset!codigovale
    If .Recordset.NoMatch = True Then
      MsgBox "Erro na tabela de Vales"
      Exit Sub
    End If
    .Recordset.Edit
    .Recordset!Cobrado = True
    .Recordset!cobradoem = Now
    .Recordset.Update
  End With
  DoEvents
  qVales.Recordset.MoveNext
Loop
With dbFormaDePG
  LucroVenda = LucroVenda + ValorBruto
  Dias = .Recordset("reembolso")
  Mes = .Recordset("mes")
  If Mes = True Then
    Intervalo = "m"
  Else
    Intervalo = "d"
  End If
  If Dias > 0 Then
    ReceberData = DateAdd(Intervalo, Dias, txtDataBordero.Value)
  Else
    Dias = .Recordset("diadomes")
    If Dias > 0 Then
      If Dias >= txtDataBordero.Day Then
        If Dias < 28 Then
          StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
        Else
          StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
          Do While IsDate(StrTemp) = False
            Dias = Dias - 1
            If Dias <= 0 Then Dias = 31
            StrTemp = Dias & "/" & (txtDataBordero.Month + 1) & "/" & txtDataBordero.Year
          Loop
        End If
        ReceberData = CDate(StrTemp)
      Else
        If Dias < 28 Then
          StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
        Else
          StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
          Do While IsDate(StrTemp) = False
            Dias = Dias - 1
            If Dias <= 0 Then Dias = 31
            StrTemp = Dias & "/" & txtDataBordero.Month & "/" & txtDataBordero.Year
          Loop
        End If
        ReceberData = CDate(StrTemp)
      End If
    End If
  End If
  If Dias > 0 Then
    Select Case Weekday(ReceberData)
      Case 1 'domingo
        ReceberData = DateAdd("d", 1, ReceberData)
      Case 7 'sábado
        ReceberData = DateAdd("d", 2, ReceberData)
    End Select
    dbCartoes.Refresh
    
    dbCartoes.Recordset.AddNew
    dbCartoes.Recordset!ValorBruto = 0
    dbCartoes.Recordset!valorliquido = 0
    dbCartoes.Recordset!CodigoConta = .Recordset("codigoconta")
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
    dbCartoes.Recordset!Conta = dbContas.Recordset("descri")
    dbCartoes.Recordset!CodigoFormaPg = .Recordset!CodigoPagamento
    dbCartoes.Recordset!Grupo = .Recordset!Grupo
    dbCartoes.Recordset!Descri = .Recordset("descri")
    dbCartoes.Recordset!DataLanc = txtDataBordero.Value
    dbCartoes.Recordset!DataPrevista = ReceberData
    dbCartoes.Recordset!ValorBruto = dbCartoes.Recordset!ValorBruto + ValorBruto
    dbCartoes.Recordset!valorliquido = dbCartoes.Recordset!valorliquido + Liquido
    If cboFuncionario.Text = "" Then
      StrTemp = qVales.Recordset!Nome
    Else
      StrTemp = cboFuncionario.Text
    End If
    dbCartoes.Recordset!Obs = StrTemp & "-" & txtObs.Text
    dbCartoes.Recordset.Update
    
  Else
    ReceberData = txtDataBordero.Value
    Select Case Weekday(ReceberData)
      Case 1 'domingo
        ReceberData = DateAdd("d", 1, ReceberData)
      Case 7 'sábado
        ReceberData = DateAdd("d", 2, ReceberData)
    End Select
    
    With dbConciliaNova
      .Recordset.AddNew
      .Recordset!CodigoConta = dbFormaDePG.Recordset("codigoconta")
      .Recordset!DataLanc = Now
      .Recordset!compensado = True
      .Recordset!Data = Date
      .Recordset!Tipo = "Vale"
      .Recordset!Codigo = 999999998
      .Recordset!Descri = Left("Total de Vales recebido de " & cboFuncionario.Text, 50)
      .Recordset!NrDocumento = Format(txtDataBordero.Value, "short date")
      .Recordset!Valor = ValorBruto
      .Recordset.Update
    End With
    dbContas.Refresh
    dbContas.Recordset.FindFirst "codigoconta=" & .Recordset("codigoconta")
    If dbContas.Recordset.NoMatch = True Then
      MsgBox "Conta " & .Recordset("contas.descri") & " não encontrada no cadastro de contas!", vbCritical, "Erro!"
    Else
      TempValor = ValorBruto
      dbContas.Recordset.Edit
      dbContas.Recordset("saldo") = dbContas.Recordset("saldo") + TempValor
      dbContas.Recordset("total") = dbContas.Recordset("saldo") + dbContas.Recordset("previsao")
      dbContas.Recordset.Update
    End If
  End If
  
End With

With dbFormaDePgRecebido
  .RecordSource = "formadepagamentorecebido2"
  .Refresh
  .Recordset.AddNew
  .Recordset("codigofechamento") = 0
  .Recordset("codigoformadepg") = dbFormaDePG.Recordset("codigoPagamento")
  .Recordset("descri") = dbFormaDePG.Recordset("descri")
  .Recordset("valorbruto") = ValorBruto
  .Recordset("valordescoper") = Operacao
  .Recordset("valordesctarifa") = Tarifa
  .Recordset("valordesconto") = DescontoPorcento
  .Recordset("valor") = Liquido
  .Recordset("operacoes") = TotalOper
  .Recordset("data") = txtDataBordero.Value
  .Recordset("hora") = Now
  .Recordset!fechamentodiario = True
  .Recordset.Update
  .Refresh
End With
Call cmdExibir_Click

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

StrOrdemVale = " order by Datacaixa, HoraIni"
StrOrdemComissao = " order by Data, HoraIni"


txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
txtDataBordero.Value = Date
With dbFuncionarios
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVales
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbVenda
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select qvendascomissoes.*, turnos.* from qvendascomissoes, turnos where qvendascomissoes.codigoturno=turnos.codigoturno"
  .Refresh
End With
With dbFormaDePG
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbFormaDePgRecebido
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesa
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qValesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendasTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaLanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With


Call cmdExibir_Click
Select Case Usuarios.Grupo.ControleVales
  Case 1 'Somente leitura
'    cmdDespesa.Enabled = False
'    cmdIncluirRecebimento.Enabled = False
'    cmdPagarComissoes.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
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
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub

