VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDespesasPagar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas a Pagar"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14190
   Icon            =   "frmDespesasPagar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   2880
      Top             =   1680
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
   Begin VB.Data dbDespesaLanc2 
      Caption         =   "dbDespesaLanc2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 order by data"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   7560
      TabIndex        =   32
      Top             =   5400
      Width           =   2055
      Begin VB.CommandButton cmdParcelar 
         Caption         =   "Parcelar"
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtValorParcelar 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker txtVencimentoParcela 
         Height          =   285
         Left            =   120
         TabIndex        =   30
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   128516097
         CurrentDate     =   37651
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Vencimento:"
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   885
      End
      Begin VB.Label Label10 
         Caption         =   "Valor:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdConfirmaBoleto 
      Caption         =   "Confirma Rec. de Boleto"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmdProrrogar 
      Caption         =   "Prorrogar"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Data dbMovimentacao 
      Caption         =   "dbMovimentacao"
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
      RecordSource    =   "select *from movimentacao"
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmDespesasPagar.frx":0442
      Height          =   1335
      Left            =   120
      OleObjectBlob   =   "frmDespesasPagar.frx":045D
      TabIndex        =   17
      Top             =   5520
      Width           =   6975
   End
   Begin VB.Data dbPagamentos 
      Caption         =   "dbPagamentos"
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
      RecordSource    =   "select *from QConciliaNovaContas where tipo='Despesa' and codigo=0"
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data QTotalData 
      Caption         =   "QTotalData"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select sum(valor - valorpago) as total from despesaslanc2 where compensado=0"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data QTotal 
      Caption         =   "QTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Select sum(valor - valorpago) as total from despesaslanc2 where compensado=0"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBCtls.DBCombo cboConta 
      Bindings        =   "frmDespesasPagar.frx":11B4
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   4560
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from contas order by descri"
      Top             =   1320
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
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 order by data"
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data dbConciliaNova 
      Caption         =   "dbConciliaNova"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from concilianova"
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   1920
      Picture         =   "frmDespesasPagar.frx":11CB
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "Imprimir"
      Top             =   6960
      Width           =   735
   End
   Begin VB.CommandButton cmdAtualiza 
      Caption         =   "Atualizar"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtNrDoc 
      Height          =   285
      Left            =   3840
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdFecharDespesa 
      Caption         =   "Finalizar Despesa"
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   6000
      TabIndex        =   26
      Top             =   6960
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   128516097
      CurrentDate     =   37651
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmDespesasPagar.frx":1C4D
      Height          =   3615
      Left            =   120
      OleObjectBlob   =   "frmDespesasPagar.frx":1C69
      TabIndex        =   0
      Top             =   120
      Width           =   13815
   End
   Begin MSComCtl2.DTPicker txtProrrogar 
      Height          =   285
      Left            =   3000
      TabIndex        =   24
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   128516097
      CurrentDate     =   37651
   End
   Begin MSComCtl2.DTPicker txtDataRecebidoBoleto 
      Height          =   285
      Left            =   7200
      TabIndex        =   19
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   128516097
      CurrentDate     =   37651
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Data de Recebimento do Boleto:"
      Height          =   195
      Left            =   7200
      TabIndex        =   18
      Top             =   3960
      Width           =   2325
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Vencimento:"
      Height          =   195
      Left            =   3000
      TabIndex        =   23
      Top             =   6960
      Width           =   885
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Total na Data:"
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label lblTotalData 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   195
      Left            =   4200
      TabIndex        =   3
      Top             =   3960
      Width           =   825
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
      Left            =   5160
      TabIndex        =   4
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblTotalPago 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Pago:"
      Height          =   195
      Left            =   3720
      TabIndex        =   15
      Top             =   5040
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Nr. Documento:"
      Height          =   195
      Left            =   3840
      TabIndex        =   9
      Top             =   4320
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Bom Para:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   4320
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   4320
      Width           =   465
   End
End
Attribute VB_Name = "frmDespesasPagar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Incluindo As Boolean

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

Printer.FontSize = 14
StrTemp = "Relatório de Contas a Pagar"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 14
StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10
StrTemp = "Data: " & Format(Dia, "long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Vencimento"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Observação"
Printer.CurrentX = 80
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1



End Sub

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

Private Sub cmdAtualiza_Click()
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

txtData.Value = Date

End Sub

Private Sub cmdConfirmaBoleto_Click()
If DateDiff("d", Date, txtDataRecebidoBoleto.Value) >= 1 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber boleto com data futura!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -40 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode receber boleto com data anterior a 40 dias!"
    Exit Sub
  End If
End If

With dbDespesaLanc
  If .Recordset.EOF = True Then
    MsgBox "Selecione um boleto para confirmar a data de recebimento do boleto!"
    Exit Sub
  End If
  If IsNull(.Recordset!datarecebido) = False Then
    MsgBox "A data já foi lançada!"
    Exit Sub
  End If
  .Recordset.Edit
  .Recordset!datarecebido = txtDataRecebidoBoleto.Value
  .Recordset.Update
End With
End Sub

Private Sub cmdFecharDespesa_Click()
Dim Resposta As Integer, CodigoDespesa As Double
With dbDespesaLanc
  CodigoDespesa = .Recordset!CodigoDespesaLanc
  .Refresh
  .Recordset.FindFirst "codigodespesalanc=" & CodigoDespesa
  If .Recordset.NoMatch = True Then
    MsgBox "Erro na tabela de despesas!"
    Exit Sub
  End If
  If Format(.Recordset!valorpago, "Currency") <> Format(.Recordset!Valor, "Currency") Then
    MsgBox "Efetue o pagamento correto primeiro!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja finalizar a despesa atual?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Edit
  .Recordset!compensado = True
  .Recordset.Update
  .Refresh
  .Refresh
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date, Total As Currency
With dbDespesaLanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  
  Cabeca Largura, Dia
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Sub-Total: " & Format(Total, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    
    StrTemp = .Recordset!Vencimento
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Obs
    Printer.CurrentX = 80
    Printer.Print StrTemp;
    
    Total = Total + .Recordset!Valor
    StrTemp = Format(.Recordset!Valor, "Currency")
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

Private Sub cmdIncluir_Click()
Call CboConta_LostFocus


With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= dbDespesaLanc.Recordset!Vencimento And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

If DateDiff("d", Date, txtData.Value) >= 40 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode pagar com data futura acima de 40 dias!"
    Exit Sub
  End If
End If
If DateDiff("d", Date, txtData.Value) <= -40 Then
  If Usuarios.Grupo.AdmEstatus <> 2 Then
    MsgBox "Somente usuário administrativo pode pagar com data anterior a 40 dias!"
    Exit Sub
  End If
End If

Incluindo = True

If cboConta.Text = "" Or cboConta.Text <> dbContas.Recordset!Descri Then
  MsgBox "Escolha uma conta válida!"
  cboConta.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um valor válido!"
  txtValor.SetFocus
  Exit Sub
End If
If txtNrDoc.Text = "" Then
  txtNrDoc.Text = " "
End If
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
  .Recordset!Codigo = dbDespesaLanc.Recordset!CodigoDespesaLanc
  .Recordset!Descri = Left(dbDespesaLanc.Recordset!Descri, 25) & " - " & Right(dbDespesaLanc.Recordset!Obs, 25)
  If Trim(txtNrDoc.Text) = "" Then
    txtNrDoc.Text = Right(dbDespesaLanc.Recordset!Obs, 35)
  End If
  If Trim(txtNrDoc.Text) = "" Then
    txtNrDoc.Text = " "
  End If
  .Recordset!NrDocumento = txtNrDoc.Text
  .Recordset!Valor = -CCur(txtValor.Text)
  .Recordset.Update
End With

With dbContas
  .Recordset.Edit
  .Recordset!Saldo = .Recordset!Saldo - CCur(txtValor.Text)
  .Recordset.Update
  .Refresh
End With
With dbMovimentacao
  .Recordset.AddNew
  .Recordset!Data = Now
  .Recordset!Tipo = "Despesa Pg."
  .Recordset!CodigoConta = dbContas.Recordset!CodigoConta
  .Recordset!Conta = dbContas.Recordset!Descri
  .Recordset!Descri = Left(dbDespesaLanc.Recordset!Descri & "-" & dbDespesaLanc.Recordset!Obs, 50)
  .Recordset!Valor = -CCur(txtValor.Text)
  .Recordset!Saldo = dbContas.Recordset!Saldo
  .Recordset.Update
  .Refresh
  .Refresh
End With
Incluindo = False
With dbDespesaLanc
  A = .Recordset!CodigoDespesaLanc
  .Refresh
  .Recordset.FindFirst "codigodespesalanc=" & A
End With
DBGrid2.SetFocus
End Sub

Private Sub cmdParcelar_Click()
Dim Resposta As Integer

If IsNumeric(txtValorParcelar.Text) = False Then
  MsgBox "Informe um valor válido para a parcela!"
  txtValorParcelar.SetFocus
  Exit Sub
End If
If dbDespesaLanc.Recordset.EOF = True Then
  MsgBox "Escolha uma despesa a ser paga!"
  Exit Sub
End If

If dbDespesaLanc.Recordset!valorpago <> 0 Then
  MsgBox "Para parcelar um valor não pode haver pagamento."
  Exit Sub
End If
If dbDespesaLanc.Recordset!Valor > 0 Then
  Resposta = MsgBox("Esta é uma despesa positiva. Deseja continuar?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  If CCur(txtValorParcelar.Text) > 0 Then
    Resposta = MsgBox("O valor está com sinal positivo. Isto irá aumentar o valor atual. Deseja continuar?", vbYesNo + vbDefaultButton2)
    If Resposta = vbNo Then Exit Sub
  End If
End If
Resposta = MsgBox("Deseja parcelar a despesa selecionada?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub


With dbDespesaLanc2
On Error Resume Next
  .Recordset.AddNew
  .Recordset!CodigoFechamento = dbDespesaLanc.Recordset!CodigoFechamento
  .Recordset!Origem = dbDespesaLanc.Recordset!Origem
  .Recordset!Data = dbDespesaLanc.Recordset!Data
  .Recordset!Hora = Now
  .Recordset!Vencimento = txtVencimentoParcela.Value
  .Recordset!CodigoDespesa = dbDespesaLanc.Recordset!CodigoDespesa
  .Recordset!NrDocumento = dbDespesaLanc.Recordset!NrDocumento
  .Recordset!Descri = dbDespesaLanc.Recordset!Descri
  .Recordset!Obs = dbDespesaLanc.Recordset!Obs
  .Recordset!Valor = -CCur(txtValorParcelar.Text)
  .Recordset!Usuario = dbDespesaLanc.Recordset!Usuario
  .Recordset!Produto = dbDespesaLanc.Recordset!Produto
  .Recordset!fechamentodiario = dbDespesaLanc.Recordset!fechamentodiario
  .Recordset!DataFechamento = dbDespesaLanc.Recordset!DataFechamento
  .Recordset.Update
End With
With dbDespesaLanc
  A = .Recordset.AbsolutePosition
  .Recordset.Edit
  .Recordset!Valor = .Recordset!Valor + CCur(txtValorParcelar.Text)
  .Recordset.Update
  .Refresh
  On Error Resume Next
  .Recordset.AbsolutePosition = A
End With

txtValorParcelar.Text = ""

End Sub

Private Sub cmdProrrogar_Click()

If Usuarios.Grupo.AdmEstatus <> 2 Then
  MsgBox "Somente usuário administrativo pode prorrogar uma despesa!"
  Exit Sub
End If

With dbDespesaLanc
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma despesa primeiro!"
    Exit Sub
  End If
  .Recordset.Edit
  .Recordset!Vencimento = txtProrrogar.Value
  CodigoDespesa = .Recordset!CodigoDespesaLanc
  .Recordset.Update
  Call cmdAtualiza_Click
  .Recordset.FindFirst "codigodespesalanc=" & CodigoDespesa
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbDespesaLanc_Reposition()
Dim CodigoDespesa As Double, TempValor As Currency

If Incluindo = True Then Exit Sub

lblTotal.Caption = Format(0, "Currency")

TempValor = 0
With dbDespesaLanc
  If .Recordset.EOF = True Then
    CodigoDespesa = 0
  Else
    On Error Resume Next
    CodigoDespesa = .Recordset!CodigoDespesaLanc
    txtProrrogar.Value = .Recordset!Vencimento
    On Error GoTo 0
  End If
  If IsNull(.Recordset!Vencimento) = True Then
    .Recordset.Edit
    .Recordset!Vencimento = .Recordset!Data
    .Recordset.Update
  End If
  
End With

With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as pago from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=" & CodigoDespesa
  .Refresh
  If IsNull(.Recordset!Pago) = False Then
    TempValor = .Recordset!Pago
  Else
    TempValor = 0
  End If
  .RecordSource = "select *from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=" & CodigoDespesa
  .Refresh
End With

With dbDespesaLanc
  If .Recordset.EOF = True Then Exit Sub
  CodigoDespesa = .Recordset!CodigoDespesaLanc
  If TempValor <> .Recordset!valorpago Then
    .Recordset.Edit
    .Recordset!valorpago = TempValor
    .Recordset.Update
  End If
'  If .Recordset!valorpago <= .Recordset!Valor Then
'    .Recordset!compensado = True
'    .Recordset.Update
'  End If
  
End With

With qTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select sum(valor - valorpago) as total from despesaslanc2 where compensado=0"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With QTotalData
  .Connect = Conectar
  .DatabaseName = Caminho
  On Error Resume Next
  .RecordSource = "Select sum(valor - valorpago) as total from despesaslanc2 where compensado=0 and vencimento=#" & DataInglesa(Trim(Str(dbDespesaLanc.Recordset!Vencimento))) & "#"
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotalData.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalData.Caption = Format(0, "Currency")
  End If
  On Error GoTo 0
End With
On Error Resume Next
txtData.Value = dbDespesaLanc.Recordset!Vencimento
lblTotalPago.Caption = Format(TempValor, "Currency")
If TempValor = 0 Then
  TempValor = -dbDespesaLanc.Recordset!Valor
Else
  TempValor = dbDespesaLanc.Recordset!Valor - TempValor
End If
txtValor.Text = Format(TempValor, "currency")

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
  Case vbKeyF5
    Call cmdAtualiza_Click
End Select
End Sub

Private Sub Form_Load()
txtDataRecebidoBoleto.Value = Date
txtVencimentoParcela.Value = Date


Incluindo = False
With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=0"
  .Refresh
End With
With dbContas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbMovimentacao
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QTotalData
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With

With dbConciliaNova
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbDespesaLanc
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from despesaslanc2 where compensado=0 order by vencimento"
  .Refresh
End With
With dbDespesaLanc2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from despesaslanc2 order by vencimento"
  .Refresh
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

txtData.Value = Date

Select Case Usuarios.Grupo.ControleContasPg
  Case 1 'Somente leitura
    cmdIncluir.Enabled = False
    cmdFecharDespesa.Enabled = False
    cmdProrrogar.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyEscape
    KeyCode = 0
    Unload Me
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
  .Text = Format(.Text, "Currency")
End With
End Sub

Private Sub txtValorParcelar_GotFocus()
With txtValorParcelar
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValorParcelar_LostFocus()
With txtValorParcelar
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "currency")
End With
End Sub
