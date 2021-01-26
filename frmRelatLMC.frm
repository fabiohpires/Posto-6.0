VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatLMC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Livro de Movimentação de Combustível (LMC)"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   Icon            =   "frmRelatLMC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   3120
      Picture         =   "frmRelatLMC.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin VB.Data dbLmcBicos 
      Caption         =   "dbLmcBicos"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMCBicos"
      Top             =   2880
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMCEstoque"
      Top             =   2160
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMC"
      Top             =   1800
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LMCNotas"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdGerar 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1095
   End
   Begin VB.Data dbTanques 
      Caption         =   "dbTanques"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Tanques"
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtMes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4080
      MaxLength       =   2
      TabIndex        =   9
      Top             =   960
      Width           =   375
   End
   Begin VB.Data dataProdutos 
      Caption         =   "dataProdutos"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from produtos where combustivel=-1 order by descri"
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data dbProdutos2 
      Caption         =   "dbProdutos2"
      Connect         =   "Access"
      DatabaseName    =   "Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from produtos where combustivel=-1 order by ordemlmc"
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton cmdTermos 
      Caption         =   "Imprime Termos"
      Height          =   375
      Left            =   4560
      TabIndex        =   10
      Top             =   840
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatLMC.frx":0EC4
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frmRelatLMC.frx":0EDE
      TabIndex        =   5
      Top             =   1440
      Width           =   6135
   End
   Begin MSDBCtls.DBCombo cboProduto 
      Bindings        =   "frmRelatLMC.frx":1C25
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   3720
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   57606145
      CurrentDate     =   38286
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   57606145
      CurrentDate     =   38286
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Mês:"
      Height          =   195
      Left            =   4080
      TabIndex        =   8
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label3 
      Caption         =   "Produto:"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatLMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database, Ws As Workspace
Dim dbBicoMovimenta           As Recordset
Dim dbNotaEntra               As Recordset
Dim dbTanqueEstoqueAbertura   As Recordset
Dim dbTanqueEstoqueFecha      As Recordset
Dim dbProdutos                As Recordset
Dim dbFornecedores            As Recordset
Dim dbPosto                   As Recordset
Dim dbTurnos                  As Recordset
Dim dbAcumulado               As Recordset

Dim UltimoDia As Date, PrimeiroDia As Date, Folhas As Integer
Dim StrTemp As String, Largura As Double, MargemE As Double
Dim InicioY As Double, InicioX As Double

Dim strOrdem As String

Private Sub ImprimeInicio()
Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.FontBold = True
StrTemp = "Portaria DNC nº 26, de 13.11.1992 - DOU 16.11.1992"
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY, Largura - 1 - MargemE

Printer.FontName = "Arial"
Printer.FontSize = 16
Printer.FontBold = True
StrTemp = "LIVRO Nº " & Format(txtMes.Text, "000000")
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 3, Largura - 1 - MargemE

Printer.FontName = "Arial"
Printer.FontSize = 16
Printer.FontBold = True
StrTemp = "LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEL (LMC)"
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 5, Largura - 1 - MargemE

Printer.CurrentY = Printer.CurrentY + 5
A = Printer.CurrentY
Printer.FontName = "Arial"
Printer.FontSize = 12
Printer.FontBold = False
StrTemp = "Tanque"
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

Printer.CurrentY = A
StrTemp = "Produto"
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 45, Printer.CurrentY, Largura - 30 - MargemE

Printer.CurrentY = A
StrTemp = "Capacidade"
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
With dbTanques
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      A = Printer.CurrentY
      StrTemp = .Recordset!Tanque
      ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
      
      Printer.CurrentY = A
      StrTemp = .Recordset!Descri
      ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 45, Printer.CurrentY, Largura - 30 - MargemE
      
      Printer.CurrentY = A
      StrTemp = .Recordset!estoquefisico
      ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
      
      .Recordset.MoveNext
    Loop
  End If
End With

Printer.CurrentY = Printer.CurrentY + 5

If dbNotaEntra.RecordCount <> 0 Then
  dbNotaEntra.MoveLast
  dbNotaEntra.MoveFirst
  
  A = Printer.CurrentY
  StrTemp = "Distribuidora"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Do While dbNotaEntra.EOF = False
    StrTemp = dbNotaEntra!fornecedor
    ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
    
    dbNotaEntra.MoveNext
  Loop
  
End If

End Sub

Private Sub ImprimeIdentifica()
If IsNull(dbPosto!Nome) = False Then
  StrTemp = "Nome: " & dbPosto!Nome
Else
  StrTemp = "Nome: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY + 1, Largura - 60 - MargemE

If IsNull(dbPosto!Endereco) = False Then
  StrTemp = "Endereço: " & dbPosto!Endereco
Else
  StrTemp = "Endereço: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

A = Printer.CurrentY
If IsNull(dbPosto!municipio) = False Then
  StrTemp = "Município: " & dbPosto!municipio
Else
  StrTemp = "Município: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

Printer.CurrentY = A
If IsNull(dbPosto!Estado) = False Then
  StrTemp = "Estado: " & dbPosto!Estado
Else
  StrTemp = "Estado: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE

A = Printer.CurrentY
If IsNull(dbPosto!JUCESP) = False Then
  StrTemp = "JUCESP: " & dbPosto!JUCESP
Else
  StrTemp = "JUCESP: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

Printer.CurrentY = A
If IsNull(dbPosto!dataJUCESP) = False Then
  StrTemp = "Data: " & dbPosto!dataJUCESP
Else
  StrTemp = "Data: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE

A = Printer.CurrentY
If IsNull(dbPosto!ie) = False Then
  StrTemp = "IE: " & dbPosto!ie
Else
  StrTemp = "IE: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

Printer.CurrentY = A
If IsNull(dbPosto!CNPJ) = False Then
  StrTemp = "CNPJ: " & dbPosto!CNPJ
Else
  StrTemp = "CNPJ: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE

If IsNull(dbPosto!ccm) = False Then
  StrTemp = "CCM: " & dbPosto!ccm
Else
  StrTemp = "CCM: "
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE

End Sub
Private Sub TermosPorProduto()
Dim UltimoDia As Date, PrimeiroDia As Date, Folhas As Integer
Dim StrTemp As String, Largura As Double, MargemE As Double
Dim InicioY As Double, InicioX As Double

With txtMes
  If IsNumeric(.Text) = False Then
    MsgBox "Informe um mês válido!"
    .SetFocus
    Exit Sub
  End If
  If CInt(.Text) > 12 Then
    MsgBox "Informe um mês válido!"
    .SetFocus
    Exit Sub
  End If
End With

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

Printer.ScaleMode = vbMillimeters
Largura = 190
MargemE = 10

Set dbPosto = db.OpenRecordset("select *from Postos")
If dbPosto.RecordCount = 0 Then
  MsgBox "Erro no cadastro do posto!"
  Exit Sub
End If

If cboProduto.Text <> "" Then
  Call cboProduto_LostFocus
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 and codigoproduto=" & dataProdutos.Recordset!CodigoProduto & " order by ordemlmc")
Else
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 order by ordemlmc")
End If

If dbProdutos.RecordCount = 0 Then
  MsgBox "Erro no cadastro de produtos!"
  Exit Sub
End If

dbProdutos.MoveLast
dbProdutos.MoveFirst

Do While dbProdutos.EOF = False
  UltimoDia = CDate("01/" & Format(txtMes.Text, "00") & "/" & Year(txtDataIni.Value))
  PrimeiroDia = UltimoDia
  UltimoDia = DateAdd("d", -1, DateAdd("m", 1, UltimoDia))
  With dbTanques
    .RecordSource = "select *from tanques where codigoproduto=" & dbProdutos!CodigoProduto & " order by tanque"
    .Refresh
  End With
  'Termo de Abertura
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = True
  StrTemp = "Folha Nr: 000001"
  Printer.CurrentY = 0
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  InicioY = Printer.CurrentY + 0.5
  
  Printer.Line (MargemE, InicioY)-(Largura, InicioY)
  Printer.Line (MargemE, InicioY)-(MargemE, 250)
  Printer.Line (Largura, InicioY)-(Largura, 250)
  Printer.Line (MargemE, 250)-(Largura, 250)
  
  Printer.CurrentY = InicioY + 0.5
  
  Printer.FontName = "Arial"
  Printer.FontSize = 18
  Printer.FontBold = True
  
  Folhas = Day(UltimoDia) + 2
  
  StrTemp = "TERMO DE ABERTURA"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 10, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = "LIVRO Nr " & Format(txtMes.Text, "000000")
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 10, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = "LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEL (LMC)"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 20, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = UCase(dbProdutos!Descri)
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY, Largura - 1 - MargemE
  
  Printer.CurrentY = Printer.CurrentY + 5
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Tanque Nr."
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Capacidade"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  With dbTanques
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        A = Printer.CurrentY
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = False
        StrTemp = .Recordset!Tanque
        ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
        
        Printer.CurrentY = A
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = False
        StrTemp = .Recordset!estoquefisico
        ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
              
        .Recordset.MoveNext
      Loop
    End If
  End With
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Contém este livro " & Format(Folhas, "000000") & " folhas eletronicamente numeradas de 000001 a " & Format(Folhas, "000000") & " e servirá para o lançamento das operações do estabelecimento do contribuinte abaixo identificado:" & Chr(vbKeyReturn)
  ImprimeTextoJustificado Printer, StrTemp, AlinhaJustificado, MargemE + 20, Printer.CurrentY + 10, Largura - 40 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Nome) = False Then
    StrTemp = "Nome: " & dbPosto!Nome
  Else
    StrTemp = "Nome: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY + 15, Largura - 60 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Endereco) = False Then
    StrTemp = "Endereço: " & dbPosto!Endereco
  Else
    StrTemp = "Endereço: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!municipio) = False Then
    StrTemp = "Município: " & dbPosto!municipio
  Else
    StrTemp = "Município: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Estado) = False Then
    StrTemp = "Estado: " & dbPosto!Estado
  Else
    StrTemp = "Estado: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!JUCESP) = False Then
    StrTemp = "JUCESP: " & dbPosto!JUCESP
  Else
    StrTemp = "JUCESP: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY + 15, Largura - 30 - MargemE
  
  Printer.CurrentY = A + 15
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!dataJUCESP) = False Then
    StrTemp = "Data: " & dbPosto!dataJUCESP
  Else
    StrTemp = "Data: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!ie) = False Then
    StrTemp = "IE: " & dbPosto!ie
  Else
    StrTemp = "IE: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!CNPJ) = False Then
    StrTemp = "CNPJ: " & dbPosto!CNPJ
  Else
    StrTemp = "CNPJ: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!ccm) = False Then
    StrTemp = "CCM: " & dbPosto!ccm
  Else
    StrTemp = "CCM: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!municipio) = False Then
    StrTemp = UCase(dbPosto!municipio & ", 01 DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
  Else
    StrTemp = UCase("SÃO PAULO, 01 DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 20, Largura - 31 - MargemE
  
  Printer.CurrentY = Printer.CurrentY + 20
  Printer.Line (MargemE + 59, Printer.CurrentY)-(Largura - 1 - Margem, Printer.CurrentY)
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "(assinatura do contribuinte ou seu representante legal)"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 0.5, Largura - 31 - MargemE
  
  Printer.NewPage
  'Termo de Encerramento
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = True
  StrTemp = "Folha Nr: " & Format(Folhas, "000000")
  Printer.CurrentY = 0
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  InicioY = Printer.CurrentY + 0.5
  
  Printer.Line (MargemE, InicioY)-(Largura, InicioY)
  Printer.Line (MargemE, InicioY)-(MargemE, 250)
  Printer.Line (Largura, InicioY)-(Largura, 250)
  Printer.Line (MargemE, 250)-(Largura, 250)
  
  Printer.CurrentY = InicioY + 0.5
  
  Printer.FontName = "Arial"
  Printer.FontSize = 18
  Printer.FontBold = True
  
  Folhas = Day(UltimoDia) + 2
  
  StrTemp = "TERMO DE ENCERRAMENTO"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 10, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = "LIVRO Nr " & Format(txtMes.Text, "000000")
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 10, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = "LIVRO DE MOVIMENTAÇÃO DE COMBUSTÍVEL (LMC)"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 20, Largura - 1 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 16
  Printer.FontBold = True
  StrTemp = UCase(dbProdutos!Descri)
  ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY, Largura - 1 - MargemE
  
  Printer.CurrentY = Printer.CurrentY + 5
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Tanque Nr."
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Capacidade"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  With dbTanques
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        A = Printer.CurrentY
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = False
        StrTemp = .Recordset!Tanque
        ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
        
        Printer.CurrentY = A
        Printer.FontName = "Arial"
        Printer.FontSize = 14
        Printer.FontBold = False
        StrTemp = .Recordset!estoquefisico
        ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
              
        .Recordset.MoveNext
      Loop
    End If
  End With
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "Contém este livro " & Format(Folhas, "000000") & " folhas eletronicamente numeradas de 000001 a " & Format(Folhas, "000000") & " e serviu para o lançamento das operações do estabelecimento do contribuinte abaixo identificado:" & Chr(vbKeyReturn)
  ImprimeTextoJustificado Printer, StrTemp, AlinhaJustificado, MargemE + 20, Printer.CurrentY + 10, Largura - 40 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Nome) = False Then
    StrTemp = "Nome: " & dbPosto!Nome
  Else
    StrTemp = "Nome: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY + 15, Largura - 60 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Endereco) = False Then
    StrTemp = "Endereço: " & dbPosto!Endereco
  Else
    StrTemp = "Endereço: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!municipio) = False Then
    StrTemp = "Município: " & dbPosto!municipio
  Else
    StrTemp = "Município: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!Estado) = False Then
    StrTemp = "Estado: " & dbPosto!Estado
  Else
    StrTemp = "Estado: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!JUCESP) = False Then
    StrTemp = "JUCESP: " & dbPosto!JUCESP
  Else
    StrTemp = "JUCESP: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY + 15, Largura - 30 - MargemE
  
  Printer.CurrentY = A + 15
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!dataJUCESP) = False Then
    StrTemp = "Data: " & dbPosto!dataJUCESP
  Else
    StrTemp = "Data: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  A = Printer.CurrentY
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!ie) = False Then
    StrTemp = "IE: " & dbPosto!ie
  Else
    StrTemp = "IE: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.CurrentY = A
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!CNPJ) = False Then
    StrTemp = "CNPJ: " & dbPosto!CNPJ
  Else
    StrTemp = "CNPJ: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 100, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!ccm) = False Then
    StrTemp = "CCM: " & dbPosto!ccm
  Else
    StrTemp = "CCM: "
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaEsquerda, MargemE + 30, Printer.CurrentY, Largura - 30 - MargemE
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  If IsNull(dbPosto!municipio) = False Then
    StrTemp = UCase(dbPosto!municipio & ", " & Format(Day(UltimoDia), "00") & " DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
  Else
    StrTemp = UCase("SÃO PAULO, " & Format(Day(UltimoDia), "00") & " DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
  End If
  ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 20, Largura - 31 - MargemE
  
  Printer.CurrentY = Printer.CurrentY + 20
  Printer.Line (MargemE + 59, Printer.CurrentY)-(Largura - 1 - Margem, Printer.CurrentY)
  
  Printer.FontName = "Arial"
  Printer.FontSize = 14
  Printer.FontBold = False
  StrTemp = "(assinatura do contribuinte ou seu representante legal)"
  ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 0.5, Largura - 31 - MargemE
  
  dbProdutos.MoveNext
  If dbProdutos.EOF = False Then
    Printer.NewPage
  Else
    Printer.EndDoc
  End If
Loop

NaoImprime:
Printer.EndDoc


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

Private Sub cmdEditar_Click()
frmLMCEdicao.Show
End Sub

Private Sub cmdGerar_Click()
Dim DiaAtual As Date, Dias As Integer
Dim EstoqueAbertura As Double
Dim PrimeiroDia As Date, Bico As Integer
Dim Abertura As Double, Encerrante As Double
Dim Acumulado As Currency, PrecoVenda As Double
Dim Desconto As Currency, VendaBico As Double, Afericao As Double
Dim CodigoLMC As Double, DiasNoMes As Double
Dim Vendas As Double



Set dbPosto = db.OpenRecordset("select *from Postos")
If cboProduto.Text <> "" Then
  Call cboProduto_LostFocus
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 and codigoproduto=" & dataProdutos.Recordset!CodigoProduto & " order by ordemlmc")
Else
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 order by ordemlmc")
End If
Set dbTurnos = db.OpenRecordset("select *from turnos order by horaini")


If dbProdutos.RecordCount = 0 Then
  MsgBox "Não existe produto cadastrado como combustível!"
  Exit Sub
End If

With dbLmc
  If cboProduto.Text <> "" Then
    .RecordSource = "select *from lmc where dia between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and codcombustivel=" & dbProdutos!CodigoProduto
  Else
    .RecordSource = "select *from lmc where dia between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  End If
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Resposta = MsgBox("Já existe LMC gerado neste período! Deseja gerar novamente?", vbYesNo + vbDefaultButton2)
    If Resposta = vbNo Then Exit Sub
  End If
End With
Dias = DateDiff("d", txtDataIni.Value, txtDataFim.Value)

Largura = 195
MargemE = 10

ProgressBar1.Max = (Dias + 1 * dbProdutos.RecordCount) + 1.0002
ProgressBar1.Value = ProgressBar1.Min
ProgressBar1.Visible = True
Do While dbProdutos.EOF = False
  ProgressBar1.Value = ProgressBar1.Min
  DiaAtual = txtDataIni.Value
  Do While DiaAtual <= txtDataFim.Value
    DoEvents
    With dbLmc
      .Recordset.FindFirst "dia=#" & DataInglesa(DiaAtual) & "# and codcombustivel=" & dbProdutos!CodigoProduto
      If .Recordset.NoMatch = True Then
        .Recordset.AddNew
      Else
        .Recordset.Edit
      End If
      CodigoLMC = .Recordset!CodLMC
      .Recordset!Dia = DiaAtual
      DiasNoMes = Day(DateAdd("d", -1, DateAdd("m", 1, "01/" & Month(DiaAtual) & "/" & Year(DiaAtual))))
      .Recordset!folha = Day(DiaAtual) + (DiasNoMes * (dbProdutos!ordemlmc - 1)) + 1
      .Recordset!descricombustivel = dbProdutos!Descri
      .Recordset!Codcombustivel = dbProdutos!CodigoProduto
      .Recordset.Update
    End With
    
    With dbLmcEstoque
      .RecordSource = "select *from lmcestoque where codlmc=" & CodigoLMC & " order by tanque"
      .Refresh
      If dbPosto!MedetanqueAntes = True Then
        dbTurnos.MoveFirst
        Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", 1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
        If dbTanqueEstoqueAbertura.RecordCount = 0 Then
          Do While dbTanqueEstoqueAbertura.RecordCount = 0
            dbTurnos.MoveNext
            If dbTurnos.EOF = False Then
              Set dbTanqueEstoqueAbertura = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
              Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", 1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
            Else
              Exit Do
            End If
          Loop
        End If
      Else
        dbTurnos.MoveLast
        Set dbTanqueEstoqueAbertura = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", -1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
        If dbTanqueEstoqueAbertura.RecordCount = 0 Then
          Do While dbTanqueEstoqueAbertura.RecordCount = 0
            DoEvents
            dbTurnos.MovePrevious
            If dbTurnos.BOF = False Then
              Set dbTanqueEstoqueAbertura = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", -1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
            Else
              Exit Do
            End If
          Loop
        End If
        dbTurnos.MoveLast
        Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
        If dbTanqueEstoqueFecha.RecordCount = 0 Then
          Do While dbTanqueEstoqueFecha.RecordCount = 0
            DoEvents
            dbTurnos.MovePrevious
            If dbTurnos.BOF = False Then
              Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
            Else
              Exit Do
            End If
          Loop
        End If
      End If
      
      
      EstoqueAbertura = 0
      If dbTanqueEstoqueAbertura.RecordCount <> 0 Then
        dbTanqueEstoqueAbertura.MoveLast
        dbTanqueEstoqueAbertura.MoveFirst
        Do While dbTanqueEstoqueAbertura.EOF = False
          DoEvents
          .Refresh
          If .Recordset.EOF = False Then
            .Recordset.FindFirst "tanque=" & dbTanqueEstoqueAbertura!tanquenr
            If .Recordset.NoMatch = False Then
              .Recordset.Edit
            Else
              .Recordset.AddNew
            End If
          Else
            .Recordset.AddNew
          End If
          .Recordset!CodLMC = CodigoLMC
          .Recordset!Tanque = dbTanqueEstoqueAbertura!tanquenr
          .Recordset!Abertura = dbTanqueEstoqueAbertura!Tanque
          .Recordset.Update
          dbTanqueEstoqueAbertura.MoveNext
        Loop
      End If
      
    End With
    
    
    Set dbNotaEntra = db.OpenRecordset("select *from qprodutosnotas where cancelada=0 and datanota=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " order by tanque")
    If dbNotaEntra.RecordCount <> 0 Then
      With dbLmcNotas
        .RecordSource = "select *from lmcnotas where codlmc=" & CodigoLMC & " order by tanquedescarga"
        .Refresh
        If .Recordset.RecordCount <> 0 Then
          Do While .Recordset.RecordCount <> 0
            .Recordset.Delete
            .Refresh
          Loop
        End If
        Do While dbNotaEntra.EOF = False
          If dbNotaEntra!CodigoProduto = dbProdutos!CodigoProduto Then
            .Refresh
            .Recordset.AddNew
            .Recordset!CodLMC = CodigoLMC
            .Recordset!codnota = dbNotaEntra!CodigoEntrada
            .Recordset!notanr = dbNotaEntra!NrNota
            .Recordset!datanota = dbNotaEntra!datanota
            .Recordset!Tanquedescarga = dbNotaEntra!Tanque
            .Recordset!Volume = dbNotaEntra!Quantidade
            .Recordset.Update
          End If
          dbNotaEntra.MoveNext
        Loop
      End With
    End If
    
    'Set dbBicoMovimenta = db.OpenRecordset("select *from qbicoencerrantes where datacaixa=#" & DataInglesa(DiaAtual) & "# and produtos.codigoproduto=" & dbProdutos!CodigoProduto & " order by bico, horaini")
    Set dbBicoMovimenta = db.OpenRecordset("SELECT BicoEncerrantes.*, Produtos.*, FechamentoDeCaixa.* FROM (BicoEncerrantes left JOIN Produtos ON BicoEncerrantes.CodigoProduto = Produtos.CodigoProduto)left JOIN FechamentoDeCaixa ON BicoEncerrantes.CodigoFechamento = FechamentoDeCaixa.CodigoFechamento where datacaixa=#" & DataInglesa(DiaAtual) & "# and produtos.codigoproduto=" & dbProdutos!CodigoProduto & " order by bico, horaini")
    If dbBicoMovimenta.RecordCount <> 0 Then
    With dbProdutos2
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.FindFirst "codigoproduto=" & dbBicoMovimenta!Codigo
        If .Recordset.NoMatch = False Then
          If IsNull(.Recordset!Desconto) = True Then
            .Recordset.Edit
            .Recordset!Desconto = 0
            .Recordset.Update
            Desconto = 0
          Else
            Desconto = .Recordset!Desconto
          End If
        Else
          Desconto = 0
        End If
      Else
        Desconto = 0
      End If
    End With
    End If
    If dbBicoMovimenta.RecordCount <> 0 Then
      dbBicoMovimenta.MoveLast
      dbBicoMovimenta.MoveFirst
      With dbLmcBicos
        .RecordSource = "Select *from lmcbicos where codlmc=" & CodigoLMC & " order by bico"
        .Refresh
      End With
      Bico = dbBicoMovimenta!Bico
      Tanque = dbBicoMovimenta!Tanque
      Abertura = dbBicoMovimenta!Abertura
      TotalVendaDia = 0
      Afericao = 0
      Vendas = 0
      Do While dbBicoMovimenta.EOF = False
        DoEvents
        If Bico <> dbBicoMovimenta!Bico Then
          If Abertura <> 0 And Encerrante <> 0 Then
            If Abertura > Encerrante Then
              Encerrante = Encerrante + 1000000
            End If
            With dbLmcBicos
              If .Recordset.RecordCount <> 0 Then
                .Recordset.FindFirst "bico=" & Bico
                If .Recordset.NoMatch = False Then
                  .Recordset.Edit
                Else
                  .Recordset.AddNew
                End If
              Else
                .Recordset.AddNew
              End If
              .Recordset!CodLMC = CodigoLMC
              .Recordset!Tanque = Tanque
              .Recordset!Bico = Bico
              .Recordset!Fechamento = Encerrante
              .Recordset!Abertura = Abertura
              .Recordset!afericoes = Afericao
              VendaBico = Encerrante - Abertura - Afericao
              .Recordset!Vendas = VendaBico
              .Recordset!PrecoVenda = TotalVendaDia
              .Recordset.Update
            End With
          End If
          Abertura = dbBicoMovimenta!Abertura
          Vendas = 0
          Afericao = 0
          Encerrante = 0
          Tanque = dbBicoMovimenta!Tanque
          Bico = dbBicoMovimenta!Bico
          TotalVendaDia = 0
        End If
        Vendas = Vendas + dbBicoMovimenta!Vendas
        Afericao = Afericao + dbBicoMovimenta!Retorno
        Encerrante = Abertura + Vendas + Afericao
        TotalVendaDia = TotalVendaDia + ((dbBicoMovimenta!Preco - Desconto) * dbBicoMovimenta!Vendas)
        dbBicoMovimenta.MoveNext
      Loop
      dbBicoMovimenta.MovePrevious
      If Abertura <> 0 And Encerrante <> 0 Then
        If Abertura > Encerrante Then
          Encerrante = Encerrante + 1000000
        End If
        With dbLmcBicos
          If .Recordset.EOF = False Then
            .Recordset.FindFirst "bico=" & Bico
            If .Recordset.NoMatch = False Then
              .Recordset.Edit
            Else
              .Recordset.AddNew
            End If
          Else
            .Recordset.AddNew
          End If
          .Recordset!CodLMC = CodigoLMC
          .Recordset!Tanque = Tanque
          .Recordset!Bico = Bico
          .Recordset!Fechamento = Encerrante
          .Recordset!Abertura = Abertura
          .Recordset!afericoes = Afericao
          VendaBico = Encerrante - Abertura - Afericao
          .Recordset!Vendas = VendaBico
          .Recordset!PrecoVenda = TotalVendaDia
          .Recordset.Update
        End With
      End If
    End If
    
    With dbLmcBicos
      StrTemp = .RecordSource
      .RecordSource = "select sum(precovenda) as total from lmcbicos where codlmc=" & CodigoLMC
      .Refresh
      If IsNull(.Recordset!Total) = False Then
        Total = .Recordset!Total
      Else
        Total = 0
      End If
      .RecordSource = StrTemp
      .Refresh
    End With
    dbLmc.Refresh
    If dbLmc.Recordset.EOF = False Then
      If dbLmc.Recordset!CodLMC <> CodigoLMC Then
        dbLmc.Recordset.FindFirst "codlmc=" & CodigoLMC
      End If
      dbLmc.Recordset.Edit
      dbLmc.Recordset!vendasnodia = Total
      dbLmc.Recordset.Update
    End If
    If dbPosto!MedetanqueAntes = True Then
      dbTurnos.MoveFirst
      Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", 1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
      
      If dbTanqueEstoqueFecha.RecordCount = 0 Then
        Do While dbTanqueEstoqueFecha.RecordCount = 0
          dbTurnos.MoveNext
          If dbTurnos.EOF = False Then
            Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DateAdd("d", 1, DiaAtual)) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
          Else
            Exit Do
          End If
        Loop
      End If
    Else
      dbTurnos.MoveLast
      Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
      If dbTanqueEstoqueFecha.RecordCount = 0 Then
        Do While dbTanqueEstoqueFecha.RecordCount = 0
          dbTurnos.MovePrevious
          If dbTurnos.BOF = False Then
            Set dbTanqueEstoqueFecha = db.OpenRecordset("select *from qdifcombustivel where datacaixa=#" & DataInglesa(DiaAtual) & "# and codigoproduto=" & dbProdutos!CodigoProduto & " and codigoturno=" & dbTurnos!CodigoTurno & " order by tanquenr")
          Else
            Exit Do
          End If
        Loop
      End If
    End If
    
    
    EstoqueFecha = 0
    If dbTanqueEstoqueFecha.RecordCount <> 0 Then
      dbTanqueEstoqueFecha.MoveLast
      dbTanqueEstoqueFecha.MoveFirst
      
      With dbLmcEstoque
        .RecordSource = "Select *from lmcestoque where codlmc=" & CodigoLMC & " order by tanque"
        .Refresh
        
        Do While dbTanqueEstoqueFecha.EOF = False
          If .Recordset.EOF = False Then
          .Recordset.FindFirst "tanque=" & dbTanqueEstoqueFecha!tanquenr
          If .Recordset.NoMatch = False Then
            .Recordset.Edit
          Else
            .Recordset.AddNew
          End If
          Else
            .Recordset.AddNew
          End If
          .Recordset!Tanque = dbTanqueEstoqueFecha!tanquenr
          .Recordset!CodLMC = CodigoLMC
          .Recordset!Fechamento = dbTanqueEstoqueFecha!Tanque
          .Recordset.Update
          dbTanqueEstoqueFecha.MoveNext
        Loop
      End With
      
    End If
    
    DiaAtual = DateAdd("d", 1, DiaAtual)
    On Error Resume Next
    ProgressBar1.Value = ProgressBar1.Value + 1
    On Error GoTo 0
    ProgressBar1.Refresh
  Loop
  dbProdutos.MoveNext
Loop

ProgressBar1.Visible = False

End Sub

Private Sub cmdImprime_Click()
Dim DiaAtual As Date, Dias As Integer
Dim Largura As Double, StrTemp As String
Dim InicioY As Double, EstoqueAbertura As Double
Dim TotalRecebido As Double, TotalDisponivel As Double
Dim Abertura As Double, Encerrante As Double, Bico As Integer
Dim Afericao As Double, VendaBico As Double, VendaDia As Double
Dim EstoqueFecha As Double, PrimeiroDia As Date
Dim Acumulado As Currency, PrecoVenda As Double
Dim Desconto As Currency, TotalVendaDia As Currency
Dim FolhaNR As Double
Dim MargemE As Double

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

Printer.ScaleMode = vbMillimeters


Set dbPosto = db.OpenRecordset("select *from Postos")
If cboProduto.Text <> "" Then
  Call cboProduto_LostFocus
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 and codigoproduto=" & dataProdutos.Recordset!CodigoProduto & " order by ordemlmc")
Else
  Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 order by ordemlmc")
End If
Set dbTurnos = db.OpenRecordset("select *from turnos order by horaini")


If dbProdutos.RecordCount = 0 Then
  MsgBox "Não existe produto cadastrado como combustível!"
  Exit Sub
End If


Dias = DateDiff("d", txtDataIni.Value, txtDataFim.Value)

Largura = 195
MargemE = 10

ProgressBar1.Max = (Dias + 1 * dbProdutos.RecordCount) + 1.0002
ProgressBar1.Value = ProgressBar1.Min
ProgressBar1.Visible = True
Do While dbProdutos.EOF = False
  DiaAtual = txtDataIni.Value
  Do While DiaAtual <= txtDataFim.Value
    
    With dbLmc
      .RecordSource = "Select *from lmc where dia=#" & DataInglesa(DiaAtual) & "# and codcombustivel=" & dbProdutos!CodigoProduto
      .Refresh
      If .Recordset.RecordCount = 0 Then
        MsgBox "O LMC de " & dbProdutos!Descri & " do dia " & Format(DiaAtual, "short date") & " ainda não foi gerado!"
        Printer.KillDoc
        Exit Sub
      End If
    End With
    With dbLmcEstoque
      .RecordSource = "Select *from lmcestoque where codlmc=" & dbLmc.Recordset!CodLMC & " order by tanque"
      .Refresh
    End With
    With dbLmcNotas
      .RecordSource = "select *from lmcnotas where codlmc=" & dbLmc.Recordset!CodLMC & " order by notanr, datanota"
      .Refresh
    End With
    With dbLmcBicos
      .RecordSource = "select *from lmcbicos where codlmc=" & dbLmc.Recordset!CodLMC & " order by bico"
      .Refresh
    End With

    
    
    Printer.FontName = "Arial"
    Printer.FontSize = 14
    Printer.FontBold = True
    StrTemp = "Livro de Movimentação de Combustíveis (LMC)"
    Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2) + MargemE
    Printer.Print StrTemp
    
    Printer.Print ""
    
    Printer.FontSize = 8
    Printer.FontBold = True
    
    FolhaNR = dbLmc.Recordset!folha
    StrTemp = "Fl. Nr.: " & Format(FolhaNR, "000000")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Printer.FontName = "Courier new"
    Printer.FontSize = 8
    Printer.FontBold = True
    StrTemp = "Posto: "
    Printer.CurrentX = MargemE
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = dbPosto!Nome
    Printer.Print StrTemp
    
    Printer.Print ""
    
    Printer.DrawWidth = 3
    InicioY = Printer.CurrentY
    Printer.Line (MargemE, Printer.CurrentY)-(Largura, Printer.CurrentY + 234), , B
    Printer.Line (MargemE, InicioY + 8)-(Largura, InicioY + 8)
    Printer.Line (MargemE, InicioY + 30)-(Largura, InicioY + 30)
    Printer.Line (MargemE, InicioY + 61)-(113, InicioY + 61)
    Printer.Line (113, InicioY + 66)-(Largura, InicioY + 66)
    Printer.Line (113, InicioY + 61)-(113, InicioY + 66)
    Printer.Line (MargemE, InicioY + 113)-(Largura, InicioY + 113)
    Printer.Line (MargemE, InicioY + 130)-(115, InicioY + 130)
    Printer.Line (115, InicioY + 139)-(Largura, InicioY + 139)
    Printer.Line (MargemE, InicioY + 165)-(115, InicioY + 165)
    Printer.Line (115, InicioY + 169)-(Largura, InicioY + 169)
    Printer.Line (115, InicioY + 130)-(115, InicioY + 195)
    Printer.Line (MargemE, InicioY + 195)-(Largura, InicioY + 195)
    
    
    Printer.FontBold = True
    StrTemp = "1 Produto: "
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 1.5
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = dbProdutos!Descri
    Printer.CurrentX = MargemE + 22
    Printer.CurrentY = InicioY + 1.5
    Printer.Print StrTemp;
    
    Printer.FontBold = True
    StrTemp = "2 Data: "
    Printer.CurrentX = 150
    Printer.CurrentY = InicioY + 1.5
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(DiaAtual, "short date")
    Printer.CurrentX = 172
    Printer.CurrentY = InicioY + 1.5
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "3 Estoque de Abertura (Medição física no início do dia)"
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 9
    Printer.Print StrTemp
    
    Printer.Print ""
    A = 1
    B = Printer.CurrentY
    
    For i = 0 To 5
      Printer.FontBold = True
      StrTemp = "TQ"
      Printer.CurrentX = MargemE + A
      Printer.Print StrTemp;
      A = A + 25
    Next i
    
    Printer.FontBold = True
    StrTemp = "3.1 Estq.Abertura"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.CurrentY = B
    Printer.Print StrTemp;
    
    
    EstoqueAbertura = 0
    If dbLmcEstoque.Recordset.RecordCount <> 0 Then
      dbLmcEstoque.Recordset.MoveLast
      dbLmcEstoque.Recordset.MoveFirst
      A = 25
      B = Printer.CurrentY
      c = 1
      Do While dbLmcEstoque.Recordset.EOF = False
        Printer.FontBold = False
        StrTemp = Format(dbLmcEstoque.Recordset!Tanque, "000")
        Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
        Printer.CurrentY = B
        Printer.Print StrTemp
        
        If IsNull(dbLmcEstoque.Recordset!Abertura) = False Then
          EstoqueAbertura = EstoqueAbertura + dbLmcEstoque.Recordset!Abertura
        End If
        StrTemp = Format(dbLmcEstoque.Recordset!Abertura, "#,##0.00")
        Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        c = c + 1
        A = A + 25
        dbLmcEstoque.Recordset.MoveNext
      Loop
      If c < 6 Then
        Do While c <= 6
          Printer.FontBold = False
          StrTemp = ""
          Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
          Printer.CurrentY = B
          Printer.Print StrTemp
          
          StrTemp = Format(0, "#,##0.00")
          Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp
          c = c + 1
          A = A + 25
        Loop
      End If
      Printer.FontBold = False
      StrTemp = ""
      Printer.CurrentX = A - Printer.TextWidth(StrTemp)
      Printer.CurrentY = B
      Printer.Print StrTemp
      
      StrTemp = Format(EstoqueAbertura, "#,##0.00")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
    End If
    
    Printer.FontBold = True
    StrTemp = "4 Volume Recebido no dia (em litros)"
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 31
    Printer.Print StrTemp;
    
    Printer.FontBold = True
    StrTemp = "4.1 Nr. Tq Descarga"
    Printer.CurrentX = 115
    Printer.Print StrTemp;
    
    Printer.FontBold = True
    StrTemp = "4.2 Vol. Recebido"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    c = 1
    TotalRecebido = 0
    If dbLmcNotas.Recordset.RecordCount <> 0 Then
      dbLmcNotas.Recordset.MoveLast
      dbLmcNotas.Recordset.MoveFirst
      Do While dbLmcNotas.Recordset.EOF = False
        Printer.FontBold = False
        StrTemp = "Nota Fiscal Nr."
        Printer.CurrentX = MargemE + 1
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = dbLmcNotas.Recordset!notanr
        Printer.CurrentX = MargemE + 63 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = "de"
        Printer.CurrentX = MargemE + 65
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcNotas.Recordset!datanota, "short date")
        Printer.CurrentX = 110 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcNotas.Recordset!Tanquedescarga, "000")
        Printer.CurrentX = 155 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        TotalRecebido = TotalRecebido + dbLmcNotas.Recordset!Volume
        Printer.FontBold = False
        StrTemp = Format(dbLmcNotas.Recordset!Volume, "#,##0.00")
        Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        c = c + 1
        dbLmcNotas.Recordset.MoveNext
      Loop
    End If
    If c < 4 Then
      Do While c <= 4
        Printer.FontBold = False
        StrTemp = "Nota Fiscal Nr."
        Printer.CurrentX = MargemE + 1
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = "de"
        Printer.CurrentX = MargemE + 65
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(0, "#,##0.00")
        Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        c = c + 1
      Loop
    End If
    
    Printer.FontBold = True
    StrTemp = "4.3 Total Recebido"
    Printer.CurrentX = 115
    Printer.CurrentY = InicioY + 58
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(TotalRecebido, "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    
    Printer.FontBold = True
    StrTemp = "4.4 Volume Disponível"
    Printer.CurrentX = 115
    Printer.Print StrTemp;
    
    TotalDisponivel = TotalRecebido + EstoqueAbertura
    Printer.FontBold = False
    StrTemp = Format(TotalDisponivel, "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    
    VendaDia = 0
    
    Printer.FontBold = True
    StrTemp = "5 Volume Vendido no dia (em litros)"
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 62
    Printer.Print StrTemp
    
    Printer.FontBold = False
    StrTemp = "5.1 TQ"
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 70
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = "5.2 Bico"
    Printer.CurrentX = MargemE + 16
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = "5.3 + Fechamento"
    Printer.CurrentX = MargemE + 70 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = "5.4 - Abertura"
    Printer.CurrentX = MargemE + 101 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = "5.5 - Aferições"
    Printer.CurrentX = MargemE + 135 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = "5.6 = Vendas no Bico"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Afericao = 0
    
    If dbLmcBicos.Recordset.RecordCount <> 0 Then
      dbLmcBicos.Recordset.MoveLast
      dbLmcBicos.Recordset.MoveFirst
      Do While dbLmcBicos.Recordset.EOF = False
        If Abertura > Encerrante Then
          Encerrante = Encerrante + 1000000
        End If
        Printer.FontBold = False
        StrTemp = Format(dbLmcBicos.Recordset!Tanque, "000")
        Printer.CurrentX = MargemE + 1
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcBicos.Recordset!Bico, "000")
        Printer.CurrentX = MargemE + 16
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcBicos.Recordset!Fechamento, "#,##0.00")
        Printer.CurrentX = MargemE + 70 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcBicos.Recordset!Abertura, "#,##0.00")
        Printer.CurrentX = MargemE + 101 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        Printer.FontBold = False
        StrTemp = Format(dbLmcBicos.Recordset!afericoes, "#,##0.00")
        Afericao = Afericao + dbLmcBicos.Recordset!afericoes
        Printer.CurrentX = MargemE + 135 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        VendaBico = dbLmcBicos.Recordset!Fechamento - dbLmcBicos.Recordset!Abertura - dbLmcBicos.Recordset!afericoes
        VendaDia = VendaDia + VendaBico
        Printer.FontBold = False
        StrTemp = Format(VendaBico, "#,##0.00")
        Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        TotalVendaDia = TotalVendaDia + dbLmcBicos.Recordset!PrecoVenda
        dbLmcBicos.Recordset.MoveNext
      Loop
    End If
    
    Printer.FontBold = True
    StrTemp = "5.7 Vendas no Dia"
    Printer.CurrentX = 116
    Printer.CurrentY = InicioY + 115
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(VendaDia, "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "6 Estoque Escritural"
    Printer.CurrentX = 116
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(TotalDisponivel - VendaDia - Afericao, "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "10 Valor das Vendas"
    Printer.CurrentX = MargemE + 1
    Printer.CurrentY = InicioY + 115
    Printer.Print StrTemp
    
    Printer.FontBold = False
    StrTemp = "10.1 Valor das Vendas do dia"
    Printer.CurrentX = MargemE + 1
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    
    StrTemp = Format(TotalVendaDia, "currency")
    Printer.CurrentX = MargemE + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    TotalVendaDia = 0
    
    Printer.FontBold = False
    StrTemp = "10.2 Valor Acumulado do mês"
    Printer.CurrentX = MargemE + 1
    Printer.Print StrTemp;
    
    PrimeiroDia = "01/" & Month(DiaAtual) & "/" & Year(DiaAtual)
    StrTemp = Desconto
    StrTemp = Replace(StrTemp, ",", ".")
    
    Set dbAcumulado = db.OpenRecordset("select sum(vendasnodia) as total from lmc where dia between #" & Format(Month(dbLmc.Recordset!Dia), "00") & "/01/" & Format(dbLmc.Recordset!Dia, "YYYY") & "# and #" & DataInglesa(dbLmc.Recordset!Dia) & "# and codcombustivel=" & dbLmc.Recordset!Codcombustivel)
    
    If IsNull(dbAcumulado!Total) = False Then
      Acumulado = dbAcumulado!Total
    Else
      Acumulado = 0
    End If
    Printer.FontBold = False
    StrTemp = Format(Acumulado, "currency")
    Printer.CurrentX = MargemE + 95 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    
    
    Printer.FontBold = True
    StrTemp = "Conciliação dos Estoques"
    Printer.CurrentX = MargemE + (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.CurrentY = InicioY + 197
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "9 Fechamento de Tanque físico"
    Printer.CurrentX = MargemE + 1
    Printer.Print StrTemp
    
    Printer.Print ""
    A = 1
    B = Printer.CurrentY
    
    For i = 0 To 5
      Printer.FontBold = True
      StrTemp = "TQ"
      Printer.CurrentX = MargemE + A
      Printer.Print StrTemp;
      A = A + 25
    Next i
    
    Printer.FontBold = True
    StrTemp = "9.1 Total"
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.CurrentY = B
    Printer.Print StrTemp;
    
    
    
    EstoqueFecha = 0
    If dbLmcEstoque.Recordset.RecordCount <> 0 Then
      dbLmcEstoque.Recordset.MoveLast
      dbLmcEstoque.Recordset.MoveFirst
      A = 25
      B = Printer.CurrentY
      c = 1
      Do While dbLmcEstoque.Recordset.EOF = False
        Printer.FontBold = False
        StrTemp = Format(dbLmcEstoque.Recordset!Tanque, "000")
        Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
        Printer.CurrentY = B
        Printer.Print StrTemp
        
        If IsNull(dbLmcEstoque.Recordset!Fechamento) = False Then
          EstoqueFecha = EstoqueFecha + dbLmcEstoque.Recordset!Fechamento
        End If
        StrTemp = Format(dbLmcEstoque.Recordset!Fechamento, "#,##0.00")
        Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        c = c + 1
        A = A + 25
        dbLmcEstoque.Recordset.MoveNext
      Loop
      If c < 6 Then
        Do While c <= 6
          Printer.FontBold = False
          StrTemp = ""
          Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
          Printer.CurrentY = B
          Printer.Print StrTemp
          
          StrTemp = Format(0, "#,##0.00")
          Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
          Printer.Print StrTemp
          c = c + 1
          A = A + 25
        Loop
      End If
      Printer.FontBold = False
      StrTemp = ""
      Printer.CurrentX = MargemE + A - Printer.TextWidth(StrTemp)
      Printer.CurrentY = B
      Printer.Print StrTemp
      
      
      StrTemp = Format(EstoqueFecha, "#,##0.00")
      Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
    End If
    
    Printer.FontBold = True
    StrTemp = "7 Estoque Fechamento"
    Printer.CurrentY = InicioY + 115
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 116
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(EstoqueFecha, "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "8 - Perdas / + Sobras"
    Printer.CurrentX = 116
    Printer.Print StrTemp;
    
    Printer.FontBold = False
    StrTemp = Format(EstoqueFecha - (TotalDisponivel - VendaDia), "#,##0.00")
    Printer.CurrentX = Largura - 1 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "11 Para uso do Revendedor"
    Printer.CurrentY = InicioY + 131
    Printer.CurrentX = MargemE + 1
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "13 Observações"
    Printer.CurrentY = InicioY + 166
    Printer.CurrentX = MargemE + 1
    Printer.Print StrTemp
    
    
    If IsNull(dbLmc.Recordset!Obs) = False Then
      Printer.FontBold = False
      ImprimeTextoJustificado Printer, dbLmc.Recordset!Obs, AlinhaEsquerda, MargemE + 1, Printer.CurrentY, 100
    End If
    Printer.FontBold = True
    StrTemp = "12 Destinado a Fiscalização"
    Printer.CurrentY = InicioY + 135
    Printer.CurrentX = 116
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "DNC"
    Printer.CurrentY = InicioY + 140
    Printer.CurrentX = 116
    Printer.Print StrTemp
    
    Printer.FontBold = True
    StrTemp = "Outros Órgãos Fiscais"
    Printer.CurrentY = InicioY + 170
    Printer.CurrentX = 116
    Printer.Print StrTemp
    
    Printer.FontBold = False
    StrTemp = "(*) ATENÇÃO: SE O RESULTADO FOR NEGATIVO,"
    Printer.CurrentY = InicioY + 222
    Printer.CurrentX = MargemE + (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    Printer.FontBold = False
    StrTemp = "PODE ESTAR HAVENDO VAZAMENTO DE PRODUTO PARA O MEIO AMBIENTE"
    Printer.CurrentX = MargemE + (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
    Printer.Print StrTemp
    
    Printer.NewPage
    DiaAtual = DateAdd("d", 1, DiaAtual)
    On Error Resume Next
    ProgressBar1.Value = ProgressBar1.Value + 1
    On Error GoTo 0
    ProgressBar1.Refresh
  Loop
  dbProdutos.MoveNext
Loop
Printer.EndDoc
NaoImprime:
ProgressBar1.Visible = False
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTermos_Click()

With txtMes
  If IsNumeric(.Text) = False Then
    MsgBox "Informe um mês válido!"
    .SetFocus
    Exit Sub
  End If
  If CInt(.Text) > 12 Then
    MsgBox "Informe um mês válido!"
    .SetFocus
    Exit Sub
  End If
End With

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

Printer.ScaleMode = vbMillimeters
Largura = 190
MargemE = 10

Set dbPosto = db.OpenRecordset("select *from Postos")
If dbPosto.RecordCount = 0 Then
  MsgBox "Erro no cadastro do posto!"
  Exit Sub
End If

Set dbProdutos = db.OpenRecordset("select *from produtos where lmc=-1 order by ordemlmc")

If dbProdutos.RecordCount = 0 Then
  MsgBox "Erro no cadastro de produtos!"
  Exit Sub
End If

dbProdutos.MoveLast
dbProdutos.MoveFirst

UltimoDia = CDate("01/" & Format(txtMes.Text, "00") & "/" & Year(txtDataIni.Value))
PrimeiroDia = UltimoDia
UltimoDia = DateAdd("d", -1, DateAdd("m", 1, UltimoDia))

Set dbNotaEntra = db.OpenRecordset("select qprodutosnotas.fornecedor from qprodutosnotas where cancelada=0 and lmc=-1 and datanota between #" & DataInglesa(PrimeiroDia) & "# and #" & DataInglesa(UltimoDia) & "# and tanque<>0 group by fornecedor order by fornecedor")

With dbTanques
  .RecordSource = "select tanques.tanque, tanques.estoquefisico, produtos.descri, produtos.lmc from tanques, produtos where tanques.codigoproduto=produtos.CodigoProduto and produtos.lmc=-1 order by tanque"
  .Refresh
End With


'Termo de Abertura
Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.FontBold = True
StrTemp = "Folha Nr: 000001"
Printer.CurrentY = 0
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

InicioY = Printer.CurrentY + 0.5

Printer.Line (MargemE, InicioY)-(Largura, InicioY)
Printer.Line (MargemE, InicioY)-(MargemE, 250)
Printer.Line (Largura, InicioY)-(Largura, 250)
Printer.Line (MargemE, 250)-(Largura, 250)

Printer.CurrentY = InicioY + 0.5

Folhas = (Day(UltimoDia) * dbProdutos.RecordCount) + 2

Printer.FontName = "Arial"
Printer.FontSize = 18
Printer.FontBold = True
StrTemp = "TERMO DE ABERTURA"
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 5, Largura - 1 - MargemE

ImprimeInicio

StrTemp = "Contém este livro " & Format(Folhas, "000000") & " folhas eletronicamente numeradas de 000001 a " & Format(Folhas, "000000") & " e servirá para o lançamento das operações do estabelecimento do contribuinte abaixo identificado:" & Chr(vbKeyReturn)
ImprimeTextoJustificado Printer, StrTemp, AlinhaJustificado, MargemE + 20, Printer.CurrentY + 3, Largura - 40 - MargemE

ImprimeIdentifica

If IsNull(dbPosto!municipio) = False Then
  StrTemp = UCase(dbPosto!municipio & ", 01 DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
Else
  StrTemp = UCase("SÃO PAULO, 01 DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 20, Largura - 31 - MargemE

Printer.CurrentY = Printer.CurrentY + 20
Printer.Line (MargemE + 59, Printer.CurrentY)-(Largura - 1 - Margem, Printer.CurrentY)

StrTemp = "(assinatura do contribuinte ou seu representante legal)"
ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 0.5, Largura - 31 - MargemE

Printer.NewPage
'***********************************************************************************************************
'Termo de Encerramento
'***********************************************************************************************************

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.FontBold = True
StrTemp = "Folha Nr: " & Format(Folhas, "000000")
Printer.CurrentY = 0
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

InicioY = Printer.CurrentY + 0.5

Printer.Line (MargemE, InicioY)-(Largura, InicioY)
Printer.Line (MargemE, InicioY)-(MargemE, 250)
Printer.Line (Largura, InicioY)-(Largura, 250)
Printer.Line (MargemE, 250)-(Largura, 250)

Printer.CurrentY = InicioY + 0.5

Folhas = (Day(UltimoDia) * dbProdutos.RecordCount) + 2

Printer.FontName = "Arial"
Printer.FontSize = 18
Printer.FontBold = True
StrTemp = "TERMO DE ENCERRAMENTO"
ImprimeTextoJustificado Printer, StrTemp, AlinhaCentralizado, MargemE + 1, Printer.CurrentY + 5, Largura - 1 - MargemE

ImprimeInicio

StrTemp = "Contém este livro " & Format(Folhas, "000000") & " folhas eletronicamente numeradas de 000001 a " & Format(Folhas, "000000") & " e serviu para o lançamento das operações do estabelecimento do contribuinte abaixo identificado:" & Chr(vbKeyReturn)
ImprimeTextoJustificado Printer, StrTemp, AlinhaJustificado, MargemE + 20, Printer.CurrentY + 3, Largura - 40 - MargemE

ImprimeIdentifica

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.FontBold = False
If IsNull(dbPosto!municipio) = False Then
  StrTemp = UCase(dbPosto!municipio & ", " & Format(Day(UltimoDia), "00") & " DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
Else
  StrTemp = UCase("SÃO PAULO, " & Format(Day(UltimoDia), "00") & " DE " & Format(PrimeiroDia, "mmmm") & " DE " & Year(PrimeiroDia))
End If
ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 20, Largura - 31 - MargemE

Printer.CurrentY = Printer.CurrentY + 20
Printer.Line (MargemE + 59, Printer.CurrentY)-(Largura - 1 - Margem, Printer.CurrentY)

Printer.FontName = "Arial"
Printer.FontSize = 14
Printer.FontBold = False
StrTemp = "(assinatura do contribuinte ou seu representante legal)"
ImprimeTextoJustificado Printer, StrTemp, AlinhaDireita, MargemE + 30, Printer.CurrentY + 0.5, Largura - 31 - MargemE

NaoImprime:
Printer.EndDoc

End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
With dbProdutos2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from produtos where lmc=-1" & strOrdem
  .Refresh
End With

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
txtDataIni.Value = Date
txtDataFim.Value = Date
txtMes.Text = Format(Month(DateAdd("m", -1, Date)), "00")

Set Ws = DBEngine.Workspaces(0)
Set db = Ws.OpenDatabase(Caminho)


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
strOrdem = " order by OrdemLMC"
With dbProdutos2
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "Select *from produtos where lmc=-1" & strOrdem
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .RecordSource = "Select *from produtos where combustivel=-1" & strOrdem
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset.Edit
        .Recordset!lmc = .Recordset!Combustivel
        .Recordset.Update
        .Recordset.MoveNext
      Loop
      .RecordSource = "Select *from produtos where lmc=-1" & strOrdem
      .Refresh
    End If
  End If
End With
With dbTanques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Select Case Usuarios.Grupo.AdmLMC
  Case 1 'Somente leitura
    cmdGerar.Enabled = False
    cmdEditar.Enabled = False
    DBGrid1.AllowUpdate = False
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
