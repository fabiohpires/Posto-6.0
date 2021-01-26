VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntegrador 
   Caption         =   "Integrador p/ Posto Fácil"
   ClientHeight    =   3045
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4455
   Icon            =   "Integrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Arquivo Texto|*.txt"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   840
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   3015
      Begin MSAdodcLib.Adodc dbDestino 
         Height          =   330
         Left            =   120
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "sa"
         Password        =   "masterkey"
         RecordSource    =   ""
         Caption         =   "dbDestino"
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
      Begin MSAdodcLib.Adodc dbPagamentos 
         Height          =   330
         Left            =   120
         Top             =   600
         Width           =   2655
         _ExtentX        =   4683
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
         Connect         =   $"Integrador.frx":0442
         OLEDBString     =   $"Integrador.frx":04E2
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select *from pagamentos"
         Caption         =   "dbPagamentos"
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
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Configurar 
      Caption         =   "Configurar"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      _Version        =   393216
      FullWidth       =   281
      FullHeight      =   81
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   2400
      Top             =   840
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   1680
      Top             =   840
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4215
   End
End
Attribute VB_Name = "frmIntegrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoPosto As String, NomePosto As String
Public Caminho As String, Conectar As String, Destino As String
Public WsPrincipal As Workspace, DbPrincipal As Database
Public strMSDE As String, IniciaComWindows As Boolean
Public strDBase As String

Private Sub GravaPagamentos(ByVal Dia As Date, ByVal Turno As String, ByVal Pdv As String, ByVal CodigoPagamento As String, ByVal Descri As String, ByVal Valor As Currency)
With dbPagamentos
  .Refresh
  .Recordset.Filter = "CodigoPosto='" & CodigoPosto & "' and datacaixa=#" & Dia & "# and turno='" & Turno & "' and pdv='" & Pdv & "' and CodigoPagamento='" & CodigoPagamento & "'"
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Valor = 0
  End If
  .Recordset!CodigoPosto = CodigoPosto
  .Recordset!NomePosto = NomePosto
  .Recordset!datacaixa = Dia
  .Recordset!Turno = Turno
  .Recordset!Pdv = Pdv
  .Recordset!CodigoPagamento = CodigoPagamento
  .Recordset!Descri = Descri
  .Recordset!Valor = .Recordset!Valor + Valor
  .Recordset.Update
End With
End Sub

Private Sub ExportaPagamentos()
Dim StrTemp As String, StrLinha As String
With dbPagamentos
  .Refresh
  .Recordset.Filter = ""
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      StrTemp = Space(15)
      If IsNull(.Recordset!CodigoPagamento) = False Then
        If .Recordset!CodigoPagamento <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!CodigoPagamento) + 1) = .Recordset!CodigoPagamento
        End If
      End If
      StrLinha = "005|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Descri) = False Then
        If .Recordset!Descri <> "" Then
          Mid(StrTemp, 1) = Left(.Recordset!Descri, 15)
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Valor) = False Then
        If .Recordset!Valor <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!Valor) + 1) = .Recordset!Valor
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      With dbDestino
        .Recordset.AddNew
        .Recordset!datacaixa = dbPagamentos.Recordset!datacaixa
        .Recordset!Turno = dbPagamentos.Recordset!Turno
        .Recordset!CodigoPosto = CodigoPosto
        .Recordset!NomePosto = NomePosto
        .Recordset!linhaexportada = StrLinha
        .Recordset!planodeconta = dbPagamentos.Recordset!Pdv
        .Recordset.Update
      End With
      
      .Recordset.MoveNext
    Loop
  End If
End With






End Sub

Private Sub AtualizaAdo()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

On Error Resume Next
db.Open "Provider=" & strMSDE & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & Destino
If Err.Number <> 0 Then Exit Sub

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select planodeconta from caixas order by PlanoDeConta", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "ALTER TABLE Caixas Add PlanoDeConta nVarChar(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  Else
    db.Execute "update caixas set PlanoDeConta='2100000000'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from GruposECF", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create TABLE GruposECF (Posto nVarChar(3), Grupo nvarchar(3), Descri nvarchar(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosGrupo", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create TABLE ProdutosGrupo (Posto nVarChar(3), CodProduto bigint, Descri nvarchar(50), Grupo nvarchar(3), GrupoECF nvarchar(3))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  End If
End If
dbTemp.Close

db.Close

End Sub



Public Sub CriaDb()
Dim Catalogo As New ADOX.Catalog, Tabela As New ADOX.Table, Tabela2 As New ADOX.Table
Catalogo.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Temp.mdb"
With Tabela
    .Name = "Dados"
    .Columns.Append "CodigoPosto", adVarWChar, 10
    .Columns.Append "NomePosto", adVarWChar, 50
    .Columns.Append "LinhaExportada", adLongVarWChar, 500
    .Columns.Append "DataCaixa", adDate
    .Columns.Append "Turno", adVarWChar, 2
End With
Catalogo.Tables.Append Tabela

With Tabela2
    .Name = "Pagamentos"
    .Columns.Append "CodigoPosto", adVarWChar, 10
    .Columns.Append "NomePosto", adVarWChar, 50
    .Columns.Append "DataCaixa", adDate
    .Columns.Append "Turno", adVarWChar, 2
    .Columns.Append "PDV", adVarWChar, 15
    .Columns.Append "CodigoPagamento", adVarWChar, 15
    .Columns.Append "Descri", adVarWChar, 50
    .Columns.Append "Valor", adCurrency
End With
Catalogo.Tables.Append Tabela2

Set Catalogo = Nothing
Set Tabela = Nothing
End Sub

Public Function GravarConfiguracoes(ByVal strOrigem As String, ByVal strDestino As String, ByVal BancoDeDados As String, ByVal strCodigoPosto As String, ByVal strNomePosto As String) As Boolean
GravarConfiguracoes = False
SaveSetting App.EXEName, "Config", "Codigo", strCodigoPosto
SaveSetting App.EXEName, "Config", "Nome", strNomePosto
SaveSetting App.EXEName, "Config", "Origem", strOrigem
SaveSetting App.EXEName, "Config", "Destino", strDestino
SaveSetting App.EXEName, "Config", "MSDE", BancoDeDados

Caminho = strOrigem
Destino = strDestino
strMSDE = BancoDeDados
CodigoPosto = strCodigoPosto
NomePosto = strNomePosto


GravarConfiguracoes = True
End Function

Public Function PegaConfiguracoes() As Boolean
CodigoPosto = GetSetting(App.EXEName, "Config", "Codigo")
NomePosto = GetSetting(App.EXEName, "Config", "Nome")
Caminho = GetSetting(App.EXEName, "Config", "Origem")
Destino = GetSetting(App.EXEName, "Config", "Destino")
strMSDE = GetSetting(App.EXEName, "Config", "MSDE", "SQLOLEDB.1")

Ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador")
If Ret = "" Then
  IniciaComWindows = False
Else
  IniciaComWindows = True
End If

End Function

Private Function ImportaRegistros() As Boolean
Dim db As New ADODB.Connection, db2 As New ADODB.Connection
Dim a As Integer, StrTemp As String, Dia As Date, Turno As String
Dim Inicio As Integer, Fim As Integer, Pdv As String
Dim Contador As Double, StrLinha As String, StrTemp2 As String
Dim CodigoProduto As String, Quantidade As String, ValorVenda As String, ValorTotal As String, CodigoFuncionario As String
Dim StatusItem As String, Bico As String
Dim CodigoPagamento As String, DescriPagamento As String, ValorPagamento As Currency

On Error GoTo trataErro
  With CommonDialog1
    .FileName = Caminho
    .ShowOpen
    Caminho = .FileName
  End With

If Dir(Caminho) = "" Then
  MsgBox "Arquivo de exportação não localizado!"
  ImportaRegistros = False
  Exit Function
End If
TentaDeNovo:
a = FreeFile()
Contador = 1
Open Caminho For Input As #a

lblStatus.Caption = "Carregando Tabela de Caixas..."
lblStatus.Refresh
With dbDestino
  StrDb = "Integrador"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;password=masterkey;Initial Catalog=" & StrDb & " ;Data Source=" & Destino
  .ConnectionString = Conectar
  .RecordSource = "Select *from caixas where codigoposto='" & CodigoPosto & "' and datacaixa='" & Date & "'"
  .Refresh
End With

On Error GoTo 0
db.Open Conectar

db2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Temp.mdb;Persist Security Info=False"
db2.Execute "delete from pagamentos"
db2.Close


Procimo:
Do While EOF(a) = False
  Line Input #a, StrTemp
  If EOF(a) = False And StrTemp = "" Then
    Line Input #a, StrTemp
  End If
  Select Case Mid(StrTemp, 1, 1)
    Case "N"
      Turno = Mid(StrTemp, 3, 1)
      Dia = CDate(Mid(StrTemp, 11, 2) & "/" & Mid(StrTemp, 9, 2) & "/" & Mid(StrTemp, 5, 4))
      StrLinha = "000|" & Format(Dia, "dd/mm/yyyy") & "|" & Format(Turno, "00")
      
      Line Input #a, StrTemp
      If EOF(a) = False And StrTemp = "" Then
        Line Input #a, StrTemp
      End If
      
      If Mid(StrTemp, 1, 1) <> "X" Then
        GoTo NaoGravar
      End If
      Pdv = Mid(StrTemp, 3, 4)
      
      db.Execute "delete from caixas where codigoposto='" & CodigoPosto & "' and datacaixa='" & Dia & "' and planodeconta='" & Pdv & "'"
      
      GoTo gravaRegistro
      
    Case "C"
      If Right(StrTemp, 1) = "C" Then
        Do
          Line Input #a, StrTemp
          Line Input #a, StrTemp
        Loop While Right(StrTemp, 1) = "I"
      End If
      GoTo Procimo
    Case "I"
      'Item do cupom
      'I#7896019611749#CHOC LACTA LANCY AVELA 30G#00000001000#000001900#000000190#000000000#000000000#000000190#60#00#000000000#000000000#0001#0#00000##00011#un#00002#Loja#000000087#000000086#000000070
      
      'codigo do produto
      Inicio = 3
      Fim = InStr(Inicio, StrTemp, "#")
      CodigoProduto = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'descricao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'quantidade
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      Quantidade = Mid(StrTemp, Inicio, Fim - Inicio)
      Quantidade = CDbl(Quantidade) / 1000
      
      'valor unitario
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      ValorVenda = Mid(StrTemp, Inicio, Fim - Inicio)
      ValorVenda = CDbl(ValorVenda) / 1000
      
      'valor bruto
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'valor desconto
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'valor desconto rateado
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'valor total item
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      ValorTotal = Mid(StrTemp, Inicio, Fim - Inicio)
      ValorTotal = CDbl(ValorTotal) / 100
      
      'codigo tributacao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'aliquota
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'valor icms
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'valor iss
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'posicao do item no cupom fiscal
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'bico
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      Bico = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'encerrante venda do item
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'status do item = vazio ativo ou c=cancelado
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StatusItem = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'codigounidadevenda
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'sigla unidade venda
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'codigo setor
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'descricao setor
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'custo medio de movimentacao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'custo c/ icms movimentacao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'custo s/ icms movimentacao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      If Fim = 0 Then
        StrTemp2 = Mid(StrTemp, Inicio)
      Else
        StrTemp2 = Mid(StrTemp, Inicio, Fim - Inicio)
      End If
      
      
      
      
      StrTemp2 = Space(12)
      If Len(CodigoProduto) > 12 Then
        StrTemp2 = Right(CodigoProduto, 12)
      Else
        Mid(StrTemp2, Len(StrTemp2) - Len(CodigoProduto) + 1) = CodigoProduto
      End If
      StrLinha = "002|" & StrTemp2
      
      'bico
      StrTemp2 = Space(6)
      Mid(StrTemp2, Len(StrTemp2) - Len(Bico) + 1) = Bico
      StrLinha = StrLinha & "|" & StrTemp2
      
      'quantidade
      StrTemp2 = Space(12)
      If Quantidade <> "" Then
        Mid(StrTemp2, Len(StrTemp2) - Len(Quantidade) + 1) = Quantidade
      End If
      StrLinha = StrLinha & "|" & StrTemp2
      
      'valor da venda
      StrTemp2 = Space(12)
      If ValorVenda <> "" Then
        Mid(StrTemp2, Len(StrTemp2) - Len(ValorVenda) + 1) = ValorVenda
      End If
      StrLinha = StrLinha & "|" & StrTemp2
      
      'valor total
      StrTemp2 = Space(12)
      If ValorTotal <> "" Then
        Mid(StrTemp2, Len(StrTemp2) - Len(ValorTotal) + 1) = ValorTotal
      End If
      StrLinha = StrLinha & "|" & StrTemp2
      
      'vendedor
      StrTemp2 = Space(12)
      StrLinha = StrLinha & "|" & StrTemp2
      
      StrLinha = StrLinha & "|" & CodigoProduto
      
      GoTo gravaRegistro
    
    Case "P"
      'Pagamento
      'P#01#Dinheiro#000000190#00500#A VISTA#00#000000000000#00001#Carteira
      
      'codigo do produto
      Inicio = 3
      Fim = InStr(Inicio, StrTemp, "#")
      CodigoPagamento = Mid(StrTemp, Inicio, Fim - Inicio)
      
      
      'descricao
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      DescriPagamento = Mid(StrTemp, Inicio, Fim - Inicio)
      
      'quantidade
      Inicio = Fim + 1
      Fim = InStr(Inicio, StrTemp, "#")
      ValorPagamento = CCur(Mid(StrTemp, Inicio, Fim - Inicio))
      ValorPagamento = ValorPagamento / 100
      
       GravaPagamentos Dia, Turno, Pdv, CodigoPagamento, DescriPagamento, ValorPagamento
      
      GoTo gravaRegistro
    Case Else
      GoTo NaoGravar
  End Select
  
  
  '**********************************************************************************
  '**********************************************************************************
gravaRegistro:
  With dbDestino
    .Recordset.AddNew
    .Recordset!datacaixa = Dia
    .Recordset!Turno = Turno
    .Recordset!CodigoPosto = CodigoPosto
    .Recordset!NomePosto = NomePosto
    .Recordset!linhaexportada = StrLinha
    .Recordset!planodeconta = Pdv
    .Recordset.Update
  End With
NaoGravar:
  Contador = Contador + 1
Loop

ExportaPagamentos

trataErro:
Close #a
db.Close
ImportaRegistros = True
lblStatus.Caption = ""
End Function

Private Sub cmdExportar_Click()
If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If
Animation1.Visible = True
Animation1.Open App.Path & "\engrenagem.avi"
Animation1.Play
'*******************************************************************************************
'*******************************************************************************************
ImportaRegistros

Animation1.Visible = False
End Sub

Private Sub cmdSair_Click()
End
End Sub

Private Sub Configurar_Click()
frmConfigura.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
If Dir("Temp.mdb") = "" Then
  CriaDb
End If


'Provider=SQLNCLI.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=newtrend1
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=NEWTREND
'Provider = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Arquivos do dBASE;Initial Catalog=c:\backuppista"
frmSplash.Show
frmSplash.Refresh
frmSplash.SetFocus
On Error Resume Next
frmSplash.lblWarning.Caption = "Inicializando o sistema..."
frmSplash.lblWarning.Refresh

frmSplash.lblWarning.Caption = "Carregando configurações..."

PegaConfiguracoes

If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If

frmSplash.lblWarning.Refresh

Unload frmSplash
AtualizaAdo
With dbPagamentos
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Temp.mdb;Persist Security Info=False"
  .Refresh
End With
End Sub

