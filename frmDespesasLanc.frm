VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDespesasLanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de Contas a Pagar"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   Icon            =   "frmDespesasLanc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cboFormaDePg 
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   2055
   End
   Begin VB.ComboBox cboPagarComo 
      Height          =   315
      ItemData        =   "frmDespesasLanc.frx":0442
      Left            =   3840
      List            =   "frmDespesasLanc.frx":044C
      TabIndex        =   9
      Top             =   960
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   3960
      Top             =   4200
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
   Begin MSDBCtls.DBCombo cboSubGrupo 
      Bindings        =   "frmDespesasLanc.frx":0461
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Data dbDespesaTipoGrupo 
      Caption         =   "dbDespesaTipoGrupo"
      Connect         =   "Access"
      DatabaseName    =   "D:\Fabio\Projeto for Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from despesatiposubgrupo order by descri"
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo cboDespesa 
      Bindings        =   "frmDespesasLanc.frx":0482
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Data dbDespesaTipo 
      Caption         =   "dbDespesaTipo"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from despesatipo order by descri"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbDespesaLanc 
      Caption         =   "dbDespesaLanc"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 order by data"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CheckBox chkConfirmadas 
      Caption         =   "Exibir já pagas"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   3135
   End
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from concilianova"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbContas 
      Caption         =   "dbContas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbPagamentos 
      Caption         =   "dbPagamentos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from QConciliaNovaContas where tipo='Despesa' and codigo=0"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbMovimentacao 
      Caption         =   "dbMovimentacao"
      Connect         =   "Access 2000;"
      DatabaseName    =   "E:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from movimentacao"
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtNrDoc 
      Height          =   285
      Left            =   4680
      TabIndex        =   23
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtDiasParcelas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "&Sair"
      Height          =   375
      Left            =   8760
      TabIndex        =   30
      Top             =   5160
      Width           =   855
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   37642
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   8880
      TabIndex        =   27
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   375
      Left            =   7800
      TabIndex        =   26
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtObs 
      Height          =   285
      Left            =   3240
      TabIndex        =   1
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtVencimento 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   37642
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmDespesasLanc.frx":049E
      Height          =   315
      Left            =   2280
      TabIndex        =   21
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtVencimento2 
      Height          =   285
      Left            =   6240
      TabIndex        =   25
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   37651
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmDespesasLanc.frx":04B5
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "frmDespesasLanc.frx":04D1
      TabIndex        =   28
      Top             =   2040
      Width           =   9615
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   6480
      TabIndex        =   15
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   39470
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   8160
      TabIndex        =   17
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72941569
      CurrentDate     =   39470
   End
   Begin MSComCtl2.DTPicker txtMesAno 
      Height          =   300
      Left            =   6480
      TabIndex        =   11
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "MM/yyyy"
      Format          =   72941571
      CurrentDate     =   39470
   End
   Begin VB.TextBox txtObsAdicional 
      Height          =   285
      Left            =   6480
      TabIndex        =   13
      Top             =   960
      Width           =   3255
   End
   Begin VB.Label Label18 
      Caption         =   "Forma de Pg.:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Pagar como:"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label lblMesAno 
      AutoSize        =   -1  'True
      Caption         =   "Mês e Ano Referência:"
      Height          =   195
      Left            =   6480
      TabIndex        =   10
      Top             =   720
      Width           =   1635
   End
   Begin VB.Label lblPeriodo 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   6480
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblPeriodoA 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   7920
      TabIndex        =   16
      Top             =   960
      Width           =   90
   End
   Begin VB.Label lblSubGrupo 
      Caption         =   "Sub-Grupo:"
      Height          =   255
      Left            =   3240
      TabIndex        =   38
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   2280
      TabIndex        =   20
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
      Height          =   195
      Left            =   6240
      TabIndex        =   24
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Documento:"
      Height          =   195
      Left            =   4680
      TabIndex        =   22
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label14 
      Caption         =   "Parcelas:"
      Height          =   255
      Left            =   1560
      TabIndex        =   36
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Dias entre parcelas:"
      Height          =   255
      Left            =   2400
      TabIndex        =   37
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento:"
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   7680
      TabIndex        =   34
      Top             =   120
      Width           =   390
   End
   Begin VB.Label lblObservacoes 
      AutoSize        =   -1  'True
      Caption         =   "Observação:"
      Height          =   195
      Left            =   3240
      TabIndex        =   32
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   6480
      TabIndex        =   33
      Top             =   120
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Despesa:"
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   675
   End
   Begin VB.Label lblObsAdicional 
      AutoSize        =   -1  'True
      Caption         =   "Obs. Adicional:"
      Height          =   195
      Left            =   6480
      TabIndex        =   12
      Top             =   720
      Width           =   1065
   End
End
Attribute VB_Name = "frmDespesasLanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String, Obrigatorio As String

Private Sub cboConta_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConta_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyCode = 0
  Case vbKeyEscape
    KeyCode = 0
    Unload Me
End Select

End Sub

Private Sub CboConta_LostFocus()
Me.KeyPreview = True
With dbContas
  .Refresh
  If cboConta.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboConta.Text & "'"
  If .Recordset.NoMatch = False Then
    cboConta.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboDespesa_LostFocus()
With DbDespesaTipo
  .Refresh
  If cboDespesa.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.FindFirst "descri='" & cboDespesa.Text & "'"
  If .Recordset.NoMatch = True Then Exit Sub
  cboDespesa.Text = .Recordset("descri")
  If cboSubGrupo.Visible = True Then
    cboSubGrupo.SetFocus
  End If
  If txtObs.Visible = True Then
    txtObs.SetFocus
  End If
End With

End Sub

Private Sub chkConfirmadas_Click()
If chkConfirmadas.Value = vbChecked Then
  dbDespesalanc.RecordSource = "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 and compensado=-1 and produto=0 and codigoconta=0 and fechamento=0" & strOrdem
Else
  dbDespesalanc.RecordSource = "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 and compensado=0 and produto=0 and codigoconta=0" & strOrdem
End If
dbDespesalanc.Refresh
End Sub

Private Sub cmdIncluir_Click()
Dim CodigoDespesaLanc As Double, Parcelas As Integer, DiasParcelas As Integer
Dim Pago As Boolean, StrObs As String

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

'If DateDiff("d", Date, txtData.Value) >= 30 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar despesa com data futura acima de 30 dias!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtData.Value) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar despesa com data anterior a 15 dias!"
'    Exit Sub
'  End If
'End If
'
'If DateDiff("d", Date, txtVencimento.Value) >= 120 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar vencimento com data acima de 90 dias!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtVencimento.Value) <= -1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode lançar despesa já vencida!"
'    Exit Sub
'  End If
'End If
'If cboConta.Text <> "" Then
'  If DateDiff("d", Date, txtVencimento2.Value) >= 1 Then
'    If Usuarios.Grupo.AdmEstatus <> 2 Then
'      MsgBox "Somente usuário administrativo pode lançar pagamento com data futura!"
'      Exit Sub
'    End If
'  End If
'  If DateDiff("d", Date, txtVencimento2.Value) <= -10 Then
'    If Usuarios.Grupo.AdmEstatus <> 2 Then
'      MsgBox "Somente usuário administrativo pode lançar pagamento com data anterior a 10 dias!"
'      Exit Sub
'    End If
'  End If
'
'End If

If cboDespesa.Text <> DbDespesaTipo.Recordset("descri") Then
  MsgBox "Despesa incorreta!"
  cboDespesa.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtValor.SetFocus
  Exit Sub
End If
'If txtVencimento.Value < Date Then
'  MsgBox "Documento já vencido!"
'  Exit Sub
'End If
If IsNumeric(txtParcelas.Text) = False Then
  MsgBox "Informe um número de parcelas!"
  txtParcelas.SetFocus
  Exit Sub
End If
If CInt(txtParcelas.Text) > 1 Then
  If IsNumeric(txtDiasParcelas.Text) = False Then
    MsgBox "Informe quantos dias entre parcelas!"
    txtDiasParcelas.SetFocus
    Exit Sub
  End If
End If
If CInt(txtParcelas.Text) < 1 Then
  MsgBox "A quantidade de parcelas deve ser pelo menos 1!"
  txtParcelas.SetFocus
  Exit Sub
End If
Parcelas = CInt(txtParcelas.Text)
If Parcelas > 1 Then
  DiasParcelas = CInt(txtDiasParcelas.Text)
  If DiasParcelas < 1 Then
    MsgBox "Os dias entre parcelas deve ser maior que 0!"
    txtDiasParcelas.SetFocus
    Exit Sub
  End If
End If

If txtObs.Visible = True Then
  StrObs = txtObs.Text
Else
  StrObs = cboSubGrupo.Text
End If

Select Case Obrigatorio
  Case "Mes e Ano Referência"
    StrObs = StrObs & " - Ref. " & Format(txtMesAno.Value, "MM/yyyy")
  Case "Período"
    StrObs = StrObs & " - De: " & Format(txtDataIni.Value, "short date") & " até " & Format(txtDataFim.Value, "short date")
  Case "Obs. Adicional"
    If Trim(txtObsAdicional.Text) = "" Then
      MsgBox "É preciso incluir uma observação adicional!"
      txtObsAdicional.SetFocus
      Exit Sub
    End If
    If Len(txtObsAdicional.Text) < 5 Then
      MsgBox "Observação adicional muito curta!"
      txtObsAdicional.SetFocus
      Exit Sub
    End If
    StrObs = StrObs & " - " & txtObsAdicional.Text
End Select

If cboFormaDePg.Text = "" Then
  MsgBox "Informe uma forma de pagamento!"
  cboFormaDePg.SetFocus
  Exit Sub
End If

For i = 1 To Parcelas
  With dbDespesalanc
    .Recordset.AddNew
    A = .Recordset!CodigoDespesaLanc
    .Recordset!CodigoFechamento = 0
    .Recordset!Origem = "Despesa"
    .Recordset!Data = txtData.Value
    .Recordset!Hora = Now
    If Parcelas > 1 Then
      If i = 1 Then
        .Recordset!Vencimento = txtVencimento.Value
      Else
        .Recordset!Vencimento = DateAdd("d", (i - 1) * DiasParcelas, txtVencimento.Value)
      End If
    Else
      .Recordset!Vencimento = txtVencimento.Value
    End If
    .Recordset!CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
    .Recordset!Descri = DbDespesaTipo.Recordset!Descri
    .Recordset!Obs = StrObs
    .Recordset!Valor = -CCur(txtValor.Text) / Parcelas
    .Recordset!codigoenviar = "1"
    If cboPagarComo.Text <> "" Then
      .Recordset!pagarcomo = cboPagarComo.Text
    End If
    Pago = False
    If cboConta.Text = dbContas.Recordset!Descri Then
      Pago = True
      With dbConciliaNova
        .Recordset.AddNew
        .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
        .Recordset!DataLanc = txtData.Value
        If dbContas.Recordset!temcpmf = False Then
          .Recordset!compensado = True
          .Recordset!Data = Date
        Else
          .Recordset!compensado = False
        End If
        .Recordset!Tipo = "Despesa"
        .Recordset!Codigo = dbDespesalanc.Recordset!CodigoDespesaLanc
        .Recordset!Descri = Left(DbDespesaTipo.Recordset!Descri & " - " & StrObs, 50)
        .Recordset!NrDocumento = "333333333"
        .Recordset!Valor = dbDespesalanc.Recordset!Valor
        .Recordset.Update
      End With
      
      With dbContas
        .Recordset.Edit
        .Recordset!Saldo = .Recordset!Saldo + dbDespesalanc.Recordset!Valor
        .Recordset.Update
        .Refresh
      End With
      With dbMovimentacao
        .Recordset.AddNew
        .Recordset!Data = Now
        .Recordset!Tipo = "Despesa Pg."
        .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
        .Recordset!Conta = dbContas.Recordset!Descri
        .Recordset!Descri = Left(DbDespesaTipo.Recordset!Descri & " - " & StrObs, 50)
        .Recordset!Valor = dbDespesalanc.Recordset!Valor
        .Recordset!Saldo = dbContas.Recordset!Saldo
        .Recordset.Update
        .Refresh
        .Refresh
      End With
    End If
    
    If Pago = True Then
      .Recordset!valorpago = .Recordset!Valor
      .Recordset!compensado = True
    End If
    .Recordset!fechamentodiario = True
    .Recordset!formadepg = cboFormaDePg.Text
    .Recordset.Update
    
    .Refresh
    If .Recordset.EOF = False Then
      .Recordset.MoveLast
    End If
  End With
Next i
cboDespesa.Text = ""
txtValor.Text = ""
txtObs.Text = ""
txtParcelas.Text = ""
txtDiasParcelas.Text = ""
cboDespesa.SetFocus
cboConta.Text = ""
txtNrDoc.Text = ""
txtVencimento2.Value = Date
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbDespesalanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma despesa a ser removida!"
    Exit Sub
  End If
  If .Recordset!Produto = True Then
    MsgBox "Não é possível remover a compra de um produto!"
    Exit Sub
  End If
  If .Recordset!Fechamento = True Then
    MsgBox "Não é possível remover o registro atual!"
    Exit Sub
  End If
  If .Recordset!valorpago <> 0 Then
    MsgBox "Já existe pagamento para esta despesa!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover a despesa autal?", vbYesNo, "Remover!")
  If Resposta = vbNo Then Exit Sub
  If .Recordset!autorizacao = -1 Then
    MsgBox "Essa pendência já foi confirmada e não pode ser removida!"
    Exit Sub
  End If
  .Recordset.Delete
  .Refresh
  'DataGrid1.Refresh
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DTPicker1_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  KeyCode = 0
  SendKeys Chr(vbKeyTab)
End If
End Sub

Private Sub DTPicker1_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub dbDespesaTipo_Reposition()
Dim CodigoDespesa As Double

If DbDespesaTipo.Recordset.EOF = False Then
  CodigoDespesa = DbDespesaTipo.Recordset!CodigoDespesa
  If IsNull(DbDespesaTipo.Recordset!Obrigatorio) = False Then
    Obrigatorio = DbDespesaTipo.Recordset!Obrigatorio
  Else
    Obrigatorio = "Nenhum"
  End If
Else
  CodigoDespesa = 0
  Obrigatorio = "Nenhum"
End If
With dbDespesaTipoGrupo
  .RecordSource = "Select *from despesatiposubgrupo where codigodespesatipo=" & CodigoDespesa & " order by descri"
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount = 0 Then
    lblSubGrupo.Visible = False
    cboSubGrupo.Visible = False
    txtObs.Visible = True
    lblObservacoes.Visible = True
  Else
    lblSubGrupo.Visible = True
    cboSubGrupo.Visible = True
    txtObs.Visible = False
    lblObservacoes.Visible = False
  End If
End With
Select Case Obrigatorio
  Case "Mes e Ano Referência"
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = True
    txtMesAno.Visible = True
  Case "Período"
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = True
    lblPeriodoA.Visible = True
    txtDataIni.Visible = True
    txtDataFim.Visible = True
    lblMesAno.Visible = False
    txtMesAno.Visible = False
  Case "Obs. Adicional"
    lblObsAdicional.Visible = True
    txtObsAdicional.Visible = True
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = False
    txtMesAno.Visible = False
  Case Else
    lblObsAdicional.Visible = False
    txtObsAdicional.Visible = False
    lblPeriodo.Visible = False
    lblPeriodoA.Visible = False
    txtDataIni.Visible = False
    txtDataFim.Visible = False
    lblMesAno.Visible = False
    txtMesAno.Visible = False
End Select
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
Call chkConfirmadas_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select
End Sub

Private Sub Form_Load()
strOrdem = " order by Data"

With cboFormaDePg
  .Clear
  .AddItem "A Vista"
  .AddItem "Boleto"
End With


With dbDespesalanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from despesaslanc2 where autorizacao=0 and compensado=0 and produto=0" & strOrdem
  .Refresh
End With
With DbDespesaTipo
  .Connect = Conectar
  .DatabaseName = Caminho
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
With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=0"
  .Refresh
End With
With dbMovimentacao
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaTipoGrupo
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

txtVencimento.Value = Date
txtData.Value = Date
txtVencimento2.Value = Date
Select Case Usuarios.Grupo.ControleLancContas
  Case 1 'Somente leitura
    cmdIncluir.Enabled = False
    cmdRemover.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  KeyCode = 0
  SendKeys Chr(vbKeyTab)
End If
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

Private Sub txtMesAno_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtMesAno_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtMesAno_LostFocus()
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
With txtValor
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
End With
End Sub

Private Sub txtVencimento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtVencimento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  KeyCode = 0
  SendKeys Chr(vbKeyTab)
End If
End Sub

Private Sub txtVencimento_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtVencimento2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtVencimento2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  KeyCode = 0
  SendKeys Chr(vbKeyTab)
End If
End Sub

Private Sub txtVencimento2_LostFocus()
Me.KeyPreview = True
End Sub
