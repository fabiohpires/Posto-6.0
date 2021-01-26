VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliacaoCartao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliação de Cartões"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10560
   Icon            =   "frmConciliacaoCartao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      DataField       =   "Obs"
      DataSource      =   "dbCartoes"
      Height          =   2775
      Left            =   7920
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   3720
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   1800
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from contas"
      Caption         =   "dbContas"
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
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   1920
      Top             =   4680
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
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo cboFormaDePg 
      Bindings        =   "frmConciliacaoCartao.frx":0442
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbFormaDePg 
      Height          =   330
      Left            =   1800
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from FormaDePagamento order by descri"
      Caption         =   "dbFormaDePg"
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
   Begin VB.Data qCartoes 
      Caption         =   "qCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valorliquido) as liquido from Cartoes where confirmado=0"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data qCartoes2 
      Caption         =   "qCartoes2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Cartoes where confirmado=0 order by dataprevista"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data dbCartoes2 
      Caption         =   "dbCartoes2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Cartoes where confirmado=0 order by dataprevista"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2775
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
      Left            =   5400
      TabIndex        =   12
      Top             =   6120
      Width           =   615
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
      Left            =   6120
      TabIndex        =   13
      Top             =   6120
      Width           =   615
   End
   Begin VB.Data dbConcilia 
      Caption         =   "dbConcilia"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from ConciliaNova"
      Top             =   1440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6840
      TabIndex        =   14
      Top             =   6120
      Width           =   855
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Data dbCartoes 
      Caption         =   "dbCartoes"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from Cartoes where confirmado=0 order by dataprevista"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliacaoCartao.frx":045C
      Height          =   2775
      Left            =   120
      OleObjectBlob   =   "frmConciliacaoCartao.frx":0474
      TabIndex        =   3
      Top             =   600
      Width           =   10335
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   6240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   207290369
      CurrentDate     =   37648
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmConciliacaoCartao.frx":1537
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmConciliacaoCartao.frx":1550
      TabIndex        =   4
      Top             =   3840
      Width           =   7575
   End
   Begin VB.Label Label6 
      Caption         =   "Obs:"
      Height          =   255
      Left            =   7920
      TabIndex        =   17
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Exibir Somente:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblTotalSomado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Total Líquido:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   16
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Líquido:"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   6000
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   "Valor Rec.:"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   6000
      Width           =   855
   End
End
Attribute VB_Name = "frmConciliacaoCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim codigoSoma As String
Public CodigoConta As Double
Dim strOrdem As String

Private Sub LimpaSoma()
With dbCartoes2
  On Error Resume Next
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  Do While .Recordset.EOF = False
    .Recordset.Edit
    .Recordset!codigoSoma = " "
    .Recordset.Update
    .Recordset.MoveNext
  Loop
End With
End Sub

Public Sub Atualiza()
Dim StrTemp As String

StrTemp = ""
With dbFormaDePG
  .ConnectionString = CaminhoADO
  .Refresh
  If cboFormaDePg.Text <> "" Then
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      .Recordset.Find "descri='" & cboFormaDePg.Text & "'"
      If .Recordset.EOF = False Then
        StrTemp = " and codigoformapg=" & .Recordset!CodigoPagamento
      End If
    End If
  End If
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from Cartoes where confirmado=0 and codigoconta=" & CodigoConta & StrTemp & strOrdem
  .Refresh
End With
With dbCartoes2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from Cartoes where confirmado=0 and codigosoma='" & codigoSoma & "'" & strOrdem
  .Refresh
End With
With qCartoes
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valorliquido) as liquido from Cartoes where confirmado=0 and codigoconta=" & CodigoConta & StrTemp
  .Refresh
End With
With qCartoes2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valorliquido) as liquido from Cartoes where confirmado=0 and codigosoma='" & codigoSoma & "'"
  .Refresh
End With
With dbConcilia
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from concilianova"
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

lblTotal.Caption = ""
If IsNull(qCartoes.Recordset!Liquido) = False Then
  lblTotal.Caption = Format(qCartoes.Recordset!Liquido, "Currency")
End If
lblTotalSomado.Caption = ""
txtValor.Text = ""
If IsNull(qCartoes2.Recordset!Liquido) = False Then
  lblTotalSomado.Caption = Format(qCartoes2.Recordset!Liquido, "Currency")
  txtValor.Text = Format(qCartoes2.Recordset!Liquido, "Currency")
End If

End Sub

Private Sub Atualiza2()
With dbCartoes
  .Refresh
End With
With dbCartoes2
  .Refresh
End With
With qCartoes
  .Refresh
End With
With qCartoes2
  .Refresh
End With

lblTotal.Caption = ""
If IsNull(qCartoes.Recordset!Liquido) = False Then
  lblTotal.Caption = Format(qCartoes.Recordset!Liquido, "Currency")
End If
lblTotalSomado.Caption = ""
txtValor.Text = ""
If IsNull(qCartoes2.Recordset!Liquido) = False Then
  lblTotalSomado.Caption = Format(qCartoes2.Recordset!Liquido, "Currency")
  txtValor.Text = Format(qCartoes2.Recordset!Liquido, "Currency")
End If

End Sub


Private Sub cmdConfirmar_Click()
Dim Resposta As Integer, ValorPrevisto As Currency, ValorRecebido As Currency
Dim Diferenca As Currency, DiaRecebido As Date
Dim TempValor As Currency, FaltaReceber As Currency

cmdConfirmar.Enabled = False

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      cmdConfirmar.Enabled = True
      Exit Sub
    End If
  End If
End With

'If DateDiff("d", Date, txtData.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar cartão com data futura!"
'    cmdConfirmar.Enabled = True
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtData.Value) <= -30 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar cartão com data anterior a 10 dias!"
'    cmdConfirmar.Enabled = True
'    Exit Sub
'  End If
'End If


With dbCartoes2
  If .Recordset.EOF = True Then
    MsgBox "Selecione um registro primeiro!"
    cmdConfirmar.Enabled = True
    Exit Sub
  End If
  If IsNumeric(txtValor.Text) = False Then
    MsgBox "Informe um valor correto!"
    txtValor.SetFocus
    cmdConfirmar.Enabled = True
    Exit Sub
  End If
  Resposta = MsgBox("Deseja confirmar os cartões atuais?", vbYesNo)
  If Resposta = vbNo Then
    cmdConfirmar.Enabled = True
    Exit Sub
  End If
  If IsNumeric(lblTotalSomado.Caption) = True Then
    ValorPrevisto = CCur(lblTotalSomado.Caption)
  Else
    ValorPrevisto = 0
  End If
  FaltaReceber = CCur(txtValor.Text)
  ValorRecebido = FaltaReceber
  Diferenca = FaltaReceber - ValorPrevisto
  If Diferenca < -1 Or Diferenca > 1 Then
    Permissao = False
    frmPermissao.Show vbModal
    If Permissao = False Then
      cmdConfirmar.Enabled = True
      Exit Sub
    End If
    DiaRecebido = txtData.Value
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If .Recordset!Confirmado = False Then
      .Recordset.Edit
      .Recordset!Confirmado = True
      .Recordset!fechadiferenca = False
      If .Recordset.RecordCount = .Recordset.AbsolutePosition + 1 Then
        TempValor = FaltaReceber
        .Recordset!Diferenca = Diferenca
      Else
        TempValor = .Recordset!valorliquido
        FaltaReceber = FaltaReceber - TempValor
      End If
      .Recordset!DataRecebida = txtData.Value
      .Recordset!ValorRecebido = TempValor
      .Recordset.Update
    End If
    .Recordset.MoveNext
  Loop
End With
With dbConcilia
  .Recordset.AddNew
  .Recordset!CodigoConta = CodigoConta
  .Recordset!DataLanc = Now
  .Recordset!compensado = True
  .Recordset!Data = txtData.Value
  .Recordset!Tipo = "Cartão"
  .Recordset!Codigo = CDbl(codigoSoma)
  dbCartoes2.Recordset.MoveFirst
  .Recordset!Descri = dbCartoes2.Recordset!Descri
  .Recordset!NrDocumento = codigoSoma
  .Recordset!Valor = ValorRecebido
  .Recordset.Update
End With
With dbContas
  .Refresh
  If .Recordset.EOF = False And .Recordset.BOF = False Then
    .Recordset.MoveFirst
    .Recordset.Find "codigoconta=" & CodigoConta
    If .Recordset.EOF = False Then
      .Recordset!Saldo = .Recordset!Saldo + ValorRecebido
      .Recordset.Update
    Else
      MsgBox "Erro na tabela de Contas!"
    End If
  End If
End With
codigoSoma = GeraCodigo()
Atualiza
cmdConfirmar.Enabled = True

End Sub

Private Sub cmdExibir_Click()
Atualiza
End Sub

Private Sub cmdSair_Click()
LimpaSoma
Unload Me
End Sub

Private Sub cmdSomar_Click()
With dbCartoes
  If .Recordset.EOF = True Then Exit Sub
  A = .Recordset.AbsolutePosition
  .Recordset.Edit
  .Recordset!codigoSoma = codigoSoma
  .Recordset.Update
  txtData.Value = .Recordset!DataPrevista
  Atualiza2
  DBGrid1.SetFocus
  On Error Resume Next
  .Recordset.AbsolutePosition = A
End With
End Sub

Private Sub cmdSubtrair_Click()
With dbCartoes2
  If .Recordset.EOF = True Then Exit Sub
  A = dbCartoes.Recordset.AbsolutePosition
  .Recordset.Edit
  .Recordset!codigoSoma = " "
  .Recordset.Update
  Atualiza2
  DBGrid1.SetFocus
  dbCartoes.Recordset.AbsolutePosition = A
End With
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
Atualiza
End Sub

Private Sub DBGrid2_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DBGrid2.Columns(ColIndex).DataField
End If
Atualiza
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub Form_Load()
strOrdem = " order by DataPrevista"

codigoSoma = GeraCodigo()
Atualiza
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
LimpaSoma
End Sub

Private Sub Form_Terminate()
LimpaSoma
End Sub

Private Sub Form_Unload(Cancel As Integer)
LimpaSoma
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
  Case Asc("+")
    Call cmdSomar_Click
  Case Asc("-")
    Call cmdSubtrair_Click
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_LostFocus()
With txtValor
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
End With
End Sub
