VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmLMCEdicao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edição de LMC"
   ClientHeight    =   6375
   ClientLeft      =   225
   ClientTop       =   1485
   ClientWidth     =   10965
   Icon            =   "frmLMCEdicao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3975
      Left            =   7200
      TabIndex        =   8
      Top             =   5880
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Data dbLmc2 
         Caption         =   "dbLmc2"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMC"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcBicosPos 
         Caption         =   "dbLmcBicosPos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCBicos"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcEstoquePos 
         Caption         =   "dbLmcEstoquePos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcPos 
         Caption         =   "dbLmcPos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMC"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcNotasPos 
         Caption         =   "dbLmcNotasPos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCNotas"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcBicosAnt 
         Caption         =   "dbLmcBicosAnt"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCBicos"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcEstoqueAnt 
         Caption         =   "dbLmcEstoqueAnt"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcAnt 
         Caption         =   "dbLmcAnt"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMC"
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcNotasAnt 
         Caption         =   "dbLmcNotasAnt"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCNotas"
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data qAcumulado 
         Caption         =   "qAcumulado"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   3120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data qLmcBicos 
         Caption         =   "qLmcBicos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data qLmcNotas 
         Caption         =   "qLmcNotas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data qLmcEstoque 
         Caption         =   "qLmcEstoque"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dataProdutos 
         Caption         =   "dataProdutos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcNotas 
         Caption         =   "dbLmcNotas"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCNotas"
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmc 
         Caption         =   "dbLmc"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMC"
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcEstoque 
         Caption         =   "dbLmcEstoque"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCEstoque"
         Top             =   600
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data dbLmcBicos 
         Caption         =   "dbLmcBicos"
         Connect         =   "Access"
         DatabaseName    =   "Posto.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   240
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "LMCBicos"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin MSMask.MaskEdBox txtMesAno 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   7
      Mask            =   "99/9999"
      PromptChar      =   " "
   End
   Begin VB.TextBox txtFolha 
      Alignment       =   1  'Right Justify
      DataField       =   "Folha"
      DataSource      =   "dbLmc"
      Height          =   285
      Left            =   10080
      TabIndex        =   33
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Bindings        =   "frmLMCEdicao.frx":0442
      Height          =   4935
      Left            =   120
      OleObjectBlob   =   "frmLMCEdicao.frx":0456
      TabIndex        =   34
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   " Bicos "
      Height          =   5055
      Left            =   2280
      TabIndex        =   11
      Top             =   720
      Width           =   5055
      Begin VB.TextBox Text1 
         DataField       =   "OBS"
         DataSource      =   "dbLmc"
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   3960
         Width           =   4815
      End
      Begin MSDBGrid.DBGrid DBGrid3 
         Bindings        =   "frmLMCEdicao.frx":0C95
         Height          =   1815
         Left            =   120
         OleObjectBlob   =   "frmLMCEdicao.frx":0CAE
         TabIndex        =   4
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Observações:"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblAcumuladoMes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3225
         TabIndex        =   31
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Valor acumulado do mês:"
         Height          =   195
         Left            =   1335
         TabIndex        =   30
         Top             =   3480
         Width           =   1785
      End
      Begin VB.Label lblValorVendasDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3225
         TabIndex        =   29
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendas do dia:"
         Height          =   195
         Left            =   2055
         TabIndex        =   28
         Top             =   3240
         Width           =   1065
      End
      Begin VB.Label lblDiferenca 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   27
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "- Perdas / + Sobras"
         Height          =   195
         Left            =   1755
         TabIndex        =   26
         Top             =   2880
         Width           =   1380
      End
      Begin VB.Label lblEstoqueFechamento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estoque Fechamento:"
         Height          =   195
         Left            =   1575
         TabIndex        =   24
         Top             =   2640
         Width           =   1560
      End
      Begin VB.Label lblEstoqueEscritural 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   23
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Estoque Escritural:"
         Height          =   195
         Left            =   1815
         TabIndex        =   22
         Top             =   2400
         Width           =   1320
      End
      Begin VB.Label lblVendasDia 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Vendas no Dia:"
         Height          =   195
         Left            =   2040
         TabIndex        =   20
         Top             =   2160
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Notas "
      Height          =   2295
      Left            =   7320
      TabIndex        =   10
      Top             =   3480
      Width           =   3615
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmLMCEdicao.frx":1BA1
         Height          =   1455
         Left            =   120
         OleObjectBlob   =   "frmLMCEdicao.frx":1BBA
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblVolumeDisponivel 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   19
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Total de Notas:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Estoque "
      Height          =   2655
      Left            =   7320
      TabIndex        =   9
      Top             =   720
      Width           =   3615
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmLMCEdicao.frx":2751
         Height          =   1575
         Left            =   120
         OleObjectBlob   =   "frmLMCEdicao.frx":276C
         TabIndex        =   6
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblEstFechamento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Fechamento:"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblEstAbertura 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Abertura:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Bindings        =   "frmLMCEdicao.frx":3307
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "Mês / Ano:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblFolha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10080
      TabIndex        =   13
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Folha:"
      Height          =   255
      Left            =   9480
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmLMCEdicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Carregando As Boolean

Private Sub Total()
If dbLmc.Recordset.EOF = True Then Exit Sub
With qLmcNotas
  .RecordSource = "select sum(volume) as total from lmcnotas where codlmc=" & dbLmc.Recordset!CodLMC
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblVolumeDisponivel.Caption = Format(.Recordset!Total, "#,##0")
  Else
    lblVolumeDisponivel.Caption = Format(0, "#,##0")
  End If
End With
With qLmcEstoque
  .RecordSource = "select sum(abertura) as Abre, sum(fechamento) as fecha from lmcestoque where codlmc=" & dbLmc.Recordset!CodLMC
  .Refresh
  If IsNull(.Recordset!abre) = False Then
    lblEstAbertura.Caption = Format(.Recordset!abre, "#,###")
  Else
    lblEstAbertura.Caption = Format(0, "#,##0")
  End If
  If IsNull(.Recordset!fecha) = False Then
    lblEstFechamento.Caption = Format(.Recordset!fecha, "#,###")
  Else
    lblEstFechamento.Caption = Format(0, "#,##0")
  End If
End With
With qLmcBicos
  .RecordSource = "select sum(vendas) as venda, sum(precovenda) as total from lmcbicos where codlmc=" & dbLmc.Recordset!CodLMC
  .Refresh
  If IsNull(.Recordset!Venda) = False Then
    lblVendasDia.Caption = Format(.Recordset!Venda, "#,###")
  Else
    lblVendasDia.Caption = Format(0, "#,##0")
  End If
  If IsNull(.Recordset!Total) = False Then
    lblValorVendasDia.Caption = Format(.Recordset!Total, "currency")
  Else
    lblValorVendasDia.Caption = Format(0, "currency")
  End If
End With
With qAcumulado
  .RecordSource = "select sum(vendasnodia) as total from lmc where dia between #" & Format(Month(dbLmc.Recordset!Dia), "00") & "/01/" & Format(dbLmc.Recordset!Dia, "YYYY") & "# and #" & DataInglesa(dbLmc.Recordset!Dia) & "# and codcombustivel=" & dbLmc.Recordset!Codcombustivel
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblAcumuladoMes.Caption = Format(.Recordset!Total, "currency")
  Else
    lblAcumuladoMes.Caption = Format(0, "currency")
  End If
End With
With dbLmc2
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "codlmc=" & dbLmc.Recordset!CodLMC
    If .Recordset.NoMatch = False Then
      .Recordset.Edit
      On Error Resume Next
      .Recordset!vendasnodia = CCur(lblValorVendasDia.Caption)
      .Recordset!acumuladonomes = CCur(lblAcumuladoMes.Caption)
      .Recordset.Update
      On Error GoTo 0
    End If
  End If
End With
On Error Resume Next
TempValor = CDbl(lblEstAbertura.Caption) + CDbl(lblVolumeDisponivel.Caption) - CDbl(lblVendasDia.Caption)
lblEstoqueEscritural.Caption = Format(TempValor, "#,###")

lblEstoqueFechamento.Caption = lblEstFechamento.Caption

TempValor = CDbl(lblEstoqueFechamento.Caption) - CDbl(lblEstoqueEscritural.Caption)
lblDiferenca.Caption = Format(TempValor, "#,###")
End Sub

Private Sub cboProduto_LostFocus()
With dataProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboProduto.Text & "'"
  If .Recordset.NoMatch = False Then
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdAbrir_Click()
Dim UltimoDia As Date

If IsDate("01/" & txtMesAno.Text) = False Then
  MsgBox "Informe um Mês e Ano válido!"
  txtMesAno.SetFocus
  Exit Sub
End If
If cboProduto.Text = "" Then
  MsgBox "Selecione um combustivel!"
  cboProduto.SetFocus
  Exit Sub
End If
Call cboProduto_LostFocus
If cboProduto.Text <> dataProdutos.Recordset!Descri Then
  MsgBox "Selecione um combustivel!"
  cboProduto.SetFocus
  Exit Sub
End If

UltimoDia = DateAdd("m", 1, "01/" & txtMesAno.Text)
UltimoDia = DateAdd("d", -1, UltimoDia)

With dbLmc
  .RecordSource = "Select *from lmc where dia between #" & DataInglesa("01/" & txtMesAno.Text) & "# and #" & DataInglesa(UltimoDia) & "# and codcombustivel=" & dataProdutos.Recordset!CodigoProduto & " order by dia, folha"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "O LMC do mes e ano selecionado ainda não foi gerado!"
    Exit Sub
  End If
  On Error Resume Next
  lblFolha.Caption = .Recordset!folha
End With

txtFolha.Visible = True

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_AfterColUpdate(ByVal ColIndex As Integer)
Dim CodigoLMC As Double
CodigoLMC = dbLmc.Recordset!CodLMC
Select Case ColIndex
  Case 1
    With dbLmcEstoqueAnt
      .Refresh
      .Recordset.FindFirst "tanque=" & dbLmcEstoque.Recordset!Tanque
      If .Recordset.NoMatch = False Then
        .Recordset.Edit
        .Recordset!Fechamento = CDbl(DBGrid1.Columns(ColIndex).Text)
        .Recordset.Update
      End If
      Call cmdAbrir_Click
    End With
  Case 2
     With dbLmcEstoquePos
      .Refresh
      .Recordset.FindFirst "tanque=" & dbLmcEstoque.Recordset!Tanque
      If .Recordset.NoMatch = False Then
        .Recordset.Edit
        .Recordset!Abertura = CDbl(DBGrid1.Columns(ColIndex).Text)
        .Recordset.Update
      End If
      Call cmdAbrir_Click
    End With
End Select
dbLmc.Recordset.FindFirst "codlmc=" & CodigoLMC
End Sub

Private Sub DBGrid1_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
  Case 0
    Cancel = True
    Exit Sub
  Case 1
    With dbLmcEstoqueAnt
      .Refresh
      If .Recordset.EOF = True Then
        MsgBox "Não existe LMC anterior! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
      .Recordset.FindFirst "tanque=" & dbLmcEstoque.Recordset!Tanque
      If .Recordset.NoMatch = True Then
        MsgBox "Tanque não encontrado no LMC anterior! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
      If CDbl(DBGrid1.Columns(ColIndex).Text) < 0 Then
        MsgBox "O estoque não pode ficar negativo! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
    End With
  Case 2
     With dbLmcEstoquePos
      .Refresh
      If .Recordset.EOF = True Then
        MsgBox "Não existe LMC posterior! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
      .Recordset.FindFirst "tanque=" & dbLmcEstoque.Recordset!Tanque
      If .Recordset.NoMatch = True Then
        MsgBox "Tanque não encontrado no LMC posterior! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
      If CDbl(DBGrid1.Columns(ColIndex).Text) < 0 Then
        MsgBox "O estoque não pode ficar negativo! A alteração será cancelada!"
        Cancel = True
        Exit Sub
      End If
    End With
End Select
End Sub

Private Sub DBGrid3_AfterColUpdate(ByVal ColIndex As Integer)
Select Case ColIndex
  Case 1 'Abertura
    With dbLmcBicosAnt
      .Refresh
      .Recordset.FindFirst "bico=" & dbLmcBicos.Recordset!Bico
      .Recordset.Edit
      .Recordset!Fechamento = CDbl(DBGrid3.Columns(ColIndex).Text)
      Preco = .Recordset!PrecoVenda / .Recordset!Vendas
      DifEstoque = .Recordset!Vendas
      .Recordset!Vendas = .Recordset!Fechamento - .Recordset!Abertura - .Recordset!afericoes
      DifEstoque = DifEstoque - .Recordset!Vendas
      .Recordset!PrecoVenda = Preco * .Recordset!Vendas
      .Recordset.Update
    End With
    With dbLmcBicos
      .Recordset.Edit
      Preco = .Recordset!PrecoVenda / .Recordset!Vendas
      .Recordset!Vendas = .Recordset!Fechamento - CDbl(DBGrid3.Columns(ColIndex).Text) - .Recordset!afericoes
      .Recordset!PrecoVenda = Preco * .Recordset!Vendas
      .Recordset.Update
    End With
    Total
  Case 2 'Encerrante
    With dbLmcBicosPos
      .Refresh
      .Recordset.FindFirst "bico=" & dbLmcBicos.Recordset!Bico
      .Recordset.Edit
      .Recordset!Abertura = CDbl(DBGrid3.Columns(ColIndex).Text)
      Preco = .Recordset!PrecoVenda / .Recordset!Vendas
      .Recordset!Vendas = .Recordset!Fechamento - .Recordset!Abertura - .Recordset!afericoes
      .Recordset!PrecoVenda = Preco * .Recordset!Vendas
      .Recordset.Update
    End With
    With dbLmcBicos
      If .Recordset!PrecoVenda <> 0 And .Recordset!Vendas <> 0 Then
        Preco = .Recordset!PrecoVenda / .Recordset!Vendas
      End If
      .Recordset.Edit
      .Recordset!Vendas = CDbl(DBGrid3.Columns(ColIndex).Text) - .Recordset!Abertura - .Recordset!afericoes
      .Recordset!PrecoVenda = Preco * .Recordset!Vendas
      .Recordset.Update
    End With
    Total
  Case 3 'Retorno
    With dbLmcBicos
      .Recordset.Edit
      Preco = .Recordset!PrecoVenda / .Recordset!Vendas
      .Recordset!Vendas = .Recordset!Fechamento - .Recordset!Abertura - CDbl(DBGrid3.Columns(ColIndex).Text)
      .Recordset!PrecoVenda = Preco * .Recordset!Vendas
      .Recordset.Update
    End With
    Total
End Select
End Sub

Private Sub DBGrid3_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
Select Case ColIndex
  Case 1 'Abertura
    If dbLmcAnt.Recordset.RecordCount = 0 Then
      MsgBox "Não existe LMC Anterior! Não será possível alterar a abertura!"
      Cancel = True
      Exit Sub
    Else
      With dbLmcBicosAnt
        .Refresh
        .Recordset.FindFirst "bico=" & dbLmcBicos.Recordset!Bico
        If .Recordset.NoMatch = True Then
          MsgBox "Bico não encontrado no LMC Anterior! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
        If CDbl(DBGrid3.Columns(ColIndex).Text) <= .Recordset!Abertura Then
          MsgBox "Encerrante menor que a abertura anterior! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
        If CDbl(DBGrid3.Columns(ColIndex).Text) > dbLmcBicos.Recordset!Fechamento Then
          MsgBox "Encerrante menor que a abertura! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
      End With
    End If
  Case 2 'Encerrante
    If dbLmcPos.Recordset.RecordCount = 0 Then
      MsgBox "Não existe LMC Posterior! Não será possível alterar o encerrante!"
      Cancel = True
      Exit Sub
    Else
      With dbLmcBicosPos
        .Refresh
        .Recordset.FindFirst "bico=" & dbLmcBicos.Recordset!Bico
        If .Recordset.NoMatch = True Then
          MsgBox "Bico não encontrado no LMC Posterior! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
        If CDbl(DBGrid3.Columns(ColIndex).Text) >= .Recordset!Fechamento Then
          MsgBox "Encerrante maior que o encerrante posterior! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
        If CDbl(DBGrid3.Columns(ColIndex).Text) < dbLmcBicos.Recordset!Abertura Then
          MsgBox "Encerrante maior que a abertura! A alteração será cancelada!"
          Cancel = True
          Exit Sub
        End If
      End With
    End If
  Case 3
    Exit Sub
  Case Else
    Cancel = True
End Select
End Sub


Private Sub dbLmc_Reposition()
Dim CodigoLMC As Double

If Carregando = True Then Exit Sub

If dbLmc.Recordset.EOF = True Then
  CodigoLMC = 0
Else
  CodigoLMC = dbLmc.Recordset!CodLMC
End If

With dbLmcEstoque
  .RecordSource = "Select *from lmcestoque where codlmc=" & CodigoLMC & " order by tanque"
  .Refresh
End With
With dbLmcNotas
  .RecordSource = "select *from lmcnotas where codlmc=" & CodigoLMC & " order by notanr, datanota"
  .Refresh
End With
With dbLmcBicos
  .RecordSource = "select *from lmcbicos where codlmc=" & CodigoLMC & " order by bico"
  .Refresh
End With

Total
If dbLmc.Recordset.EOF = True Then
  Dia = CDate("01/01/1905")
Else
  Dia = dbLmc.Recordset!Dia
End If
'Abre o anterior
On Error GoTo 0
On Error Resume Next
With dbLmcAnt
  .RecordSource = "Select *from lmc where codcombustivel=" & dbLmc.Recordset!Codcombustivel & " order by dia, folha"
  .Refresh
  Codigo = 0
  If dbLmcAnt.Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "dia=#" & DataInglesa(dbLmc.Recordset!Dia) & "# and codcombustivel=" & dbLmc.Recordset!Codcombustivel & " and folha='" & dbLmc.Recordset!folha & "'"
    If .Recordset.NoMatch = False Then
      .Recordset.MovePrevious
      If .Recordset.BOF = False Then
        Codigo = .Recordset!CodLMC
      End If
    End If
  End If
End With
With dbLmcEstoqueAnt
  .RecordSource = "Select *from lmcestoque where codlmc=" & Codigo & " order by tanque"
  .Refresh
End With
With dbLmcNotasAnt
  .RecordSource = "select *from lmcnotas where codlmc=" & Codigo & " order by notanr, datanota"
  .Refresh
End With
With dbLmcBicosAnt
  .RecordSource = "select *from lmcbicos where codlmc=" & Codigo & " order by bico"
  .Refresh
End With

'Abre o posterior
With dbLmcPos
  .RecordSource = "Select *from lmc where codcombustivel=" & dbLmc.Recordset!Codcombustivel & " order by dia, folha"
  .Refresh
  Codigo = 0
  If .Recordset.RecordCount <> 0 Then
    .Recordset.FindFirst "dia=#" & DataInglesa(dbLmc.Recordset!Dia) & "# and codcombustivel=" & dbLmc.Recordset!Codcombustivel & " and folha='" & dbLmc.Recordset!folha & "'"
    If .Recordset.NoMatch = False Then
      .Recordset.MoveNext
      If .Recordset.EOF = False Then
        Codigo = .Recordset!CodLMC
      End If
    End If
  End If
End With
With dbLmcEstoquePos
  .RecordSource = "Select *from lmcestoque where codlmc=" & Codigo & " order by tanque"
  .Refresh
End With
With dbLmcNotasPos
  .RecordSource = "select *from lmcnotas where codlmc=" & Codigo & " order by notanr, datanota"
  .Refresh
End With
With dbLmcBicosPos
  .RecordSource = "select *from lmcbicos where codlmc=" & Codigo & " order by bico"
  .Refresh
End With

If dbLmcBicos.Recordset.RecordCount = 0 Then
  If dbLmcBicosAnt.Recordset.RecordCount <> 0 Then
    With dbLmcBicosAnt
      .Refresh
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbLmcBicos.Recordset.AddNew
        dbLmcBicos.Recordset!CodLMC = dbLmc.Recordset!CodLMC
        dbLmcBicos.Recordset!Tanque = .Recordset!Tanque
        dbLmcBicos.Recordset!Bico = .Recordset!Bico
        dbLmcBicos.Recordset!Abertura = .Recordset!Fechamento
        dbLmcBicos.Recordset!Fechamento = .Recordset!Fechamento
        dbLmcBicos.Recordset!afericoes = 0
        dbLmcBicos.Recordset!Vendas = 0
        dbLmcBicos.Recordset!PrecoVenda = 0
        dbLmcBicos.Recordset.Update
        .Recordset.MoveNext
      Loop
    End With
  Else
    With dbLmcBicosPos
      If .Recordset.RecordCount <> 0 Then
        .Refresh
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          dbLmcBicos.Recordset.AddNew
          dbLmcBicos.Recordset!CodigoLMC = dbLmc.Recordset!CodigoLMC
          dbLmcBicos.Recordset!Tanque = .Recordset!Tanque
          dbLmcBicos.Recordset!Bico = .Recordset!Bico
          dbLmcBicos.Recordset!Abertura = .Recordset!Abertura
          dbLmcBicos.Recordset!Fechamento = .Recordset!Abertura
          dbLmcBicos.Recordset!afericoes = 0
          dbLmcBicos.Recordset!Vendas = 0
          dbLmcBicos.Recordset!PrecoVenda = 0
          dbLmcBicos.Recordset.Update
          .Recordset.MoveNext
        Loop
      End If
    End With
  End If
End If
If dbLmcEstoque.Recordset.RecordCount = 0 Then
  If dbLmcEstoqueAnt.Recordset.RecordCount <> 0 Then
    With dbLmcEstoqueAnt
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        dbLmcEstoque.Recordset.AddNew
        dbLmcEstoque.Recordset!CodLMC = dbLmc.Recordset!CodLMC
        dbLmcEstoque.Recordset!Tanque = .Recordset!Tanque
        dbLmcEstoque.Recordset!Abertura = .Recordset!Fechamento
        dbLmcEstoque.Recordset!Fechamento = .Recordset!Fechamento
        dbLmcEstoque.Recordset.Update
        .Recordset.MoveNext
      Loop
    End With
  Else
    With dbLmcEstoquePos
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        Do While .Recordset.EOF = False
          dbLmcEstoque.Recordset.AddNew
          dbLmcEstoque.Recordset!CodigoLMC = dbLmc.Recordset!CodigoLMC
          dbLmcEstoque.Recordset!Tanque = .Recordset!Tanque
          dbLmcEstoque.Recordset!Abertura = .Recordset!Abertura
          dbLmcEstoque.Recordset!Fechamento = .Recordset!Abertura
          dbLmcEstoque.Recordset.Update
          .Recordset.MoveNext
        Loop
      End If
    End With
  End If
End If

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
Carregando = True
txtMesAno.Text = Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
With dbLmc
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmc where codlmc=0"
  .Refresh
End With
With dbLmcEstoque
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcEstoque where codlmc=0"
  .Refresh
End With
With dbLmcNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcNotas where codlmc=0"
  .Refresh
End With
With dbLmcBicos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcBicos where codlmc=0"
  .Refresh
End With
With dataProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qLmcEstoque
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qLmcNotas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qLmcBicos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qAcumulado
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbLmcAnt
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmc where codlmc=0"
  .Refresh
End With
With dbLmcEstoqueAnt
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcEstoque where codlmc=0"
  .Refresh
End With
With dbLmcNotasAnt
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcNotas where codlmc=0"
  .Refresh
End With
With dbLmcBicosAnt
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcBicos where codlmc=0"
  .Refresh
End With

With dbLmcPos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmc where codlmc=0"
  .Refresh
End With
With dbLmcEstoquePos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcEstoque where codlmc=0"
  .Refresh
End With
With dbLmcNotasPos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcNotas where codlmc=0"
  .Refresh
End With
With dbLmcBicosPos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmcBicos where codlmc=0"
  .Refresh
End With
With dbLmc2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from lmc"
  .Refresh
End With

Carregando = False
End Sub

Private Sub txtMesAno_GotFocus()
With txtMesAno
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub
