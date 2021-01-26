VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatMovimentacaoProdutos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimentação de Produto"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   Icon            =   "frmRelatMovimentacaoProduto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin MSComCtl2.Animation Animation1 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      AutoPlay        =   -1  'True
      FullWidth       =   161
      FullHeight      =   25
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   9720
      Picture         =   "frmRelatMovimentacaoProduto.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Imprimir"
      Top             =   0
      Width           =   735
   End
   Begin MSDataListLib.DataCombo cboProduto 
      Bindings        =   "frmRelatMovimentacaoProduto.frx":0EC4
      Height          =   315
      Left            =   4320
      TabIndex        =   7
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   375
      Left            =   2520
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select codigoproduto, Codigo, Descri from produtos where combustivel=0 order by descri"
      Caption         =   "dbProdutos"
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
      Left            =   7800
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbProdutosEstoque 
      Height          =   375
      Left            =   2520
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   1
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"frmRelatMovimentacaoProduto.frx":0EDD
      Caption         =   "dbProdutosEstoque"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatMovimentacaoProduto.frx":0F64
      Height          =   5655
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9975
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "descri"
         Caption         =   "Produto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "DataCaixa"
         Caption         =   "DataCaixa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Turno"
         Caption         =   "Turno"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Abertura"
         Caption         =   "Abertura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Entrada"
         Caption         =   "Entrada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Saida"
         Caption         =   "Saida"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Acerto"
         Caption         =   "Acerto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Disponivel"
         Caption         =   "Disponivel"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.#"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2984,882
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   645,165
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   794,835
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   4320
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmRelatMovimentacaoProdutos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboProduto_GotFocus()
DataGrid1.Visible = False
End Sub

Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodigo.Text = .Recordset!Codigo
  End If
End With

End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String

StrTemp = "Select produtosestoque.*, produtos.descri from ProdutosEstoque, produtos where produtosestoque.codigoproduto=produtos.codigoproduto and datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
If dbProdutos.Recordset.RecordCount <> 0 Then
  If cboProduto.Text <> "" Then
    Call cboProduto_LostFocus
    If cboProduto.Text <> dbProdutos.Recordset!Descri Then
      MsgBox "Produto não encontrado!"
      Exit Sub
    End If
    StrTemp = StrTemp & " and produtosestoque.codigoproduto=" & dbProdutos.Recordset!CodigoProduto
  End If
End If
With dbProdutosEstoque
  .RecordSource = StrTemp & " order by produtosestoque.codigo, datacaixa, horaini"
  .Refresh
End With
DataGrid1.Visible = True
DataGrid1.SetFocus
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

StrTemp = "Período: " & txtDataIni.Value & " a " & txtDataFim.Value
If cboProduto.Text <> "" Then
  StrTemp = StrTemp & "    Cod.:" & txtCodigo.Text & " - Produto: " & cboProduto.Text
End If
ImprimeADOGrid DataGrid1, Printer, dbProdutosEstoque, 5, True, , 0, 6, 7, "Movimentação de Produtos - " & NomePosto, StrTemp, "Impresso em: " & Format(Now, "long date") & " - " & Format(Now, "short time")

Printer.EndDoc
NaoImprime:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyF5
    'If Usuarios.Nome = "Usuário Master" Then
      If cboProduto.Text <> "" Then
        Call cboProduto_LostFocus
        If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
        If dbProdutos.Recordset!Descri <> cboProduto.Text Then Exit Sub
        With Animation1
          .Visible = True
          .Open App.Path & "\engrenagem.avi"
          .Play
        End With
        EstoqueDesdeAData txtDataIni, dbProdutos.Recordset!CodigoProduto
        Animation1.Visible = False
      Else
        If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
        Resposta = MsgBox("Deseja processar todos os produtos?", vbYesNo + vbDefaultButton2)
        If Resposta = vbNo Then Exit Sub
        dbProdutos.Recordset.MoveFirst
        With Animation1
          .Visible = True
          .Open App.Path & "\engrenagem.avi"
          .Play
        End With
        Do While dbProdutos.Recordset.EOF = False
          EstoqueDesdeAData txtDataIni, dbProdutos.Recordset!CodigoProduto
          dbProdutos.Recordset.MoveNext
        Loop
        Animation1.Visible = False
      End If
      Call cmdExibir_Click
    'End If
End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
  Case vbKeyF5
    If cboProduto.Text <> "" Then
      Call cboProduto_LostFocus
      If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
      If dbProdutos.Recordset!Descri <> cboProduto.Text Then Exit Sub
      EstoqueDesdeAData txtDataIni, dbProdutos.Recordset!CodigoProduto
    Else
      If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
      Resposta = MsgBox("Deseja processar todos os produtos?", vbYesNo + vbDefaultButton2)
      If Resposta = vbNo Then Exit Sub
      dbProdutos.Recordset.MoveFirst
      Do While dbProdutos.Recordset.EOF = False
        EstoqueDesdeAData txtDataIni, dbProdutos.Recordset!CodigoProduto
        dbProdutos.Recordset.MoveNext
      Loop
    End If
End Select
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date

With dbProdutosEstoque
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .Refresh
End With
End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
DataGrid1.Visible = False
End Sub

Private Sub txtCodigo_LostFocus()
With txtCodigo
  If .Text = "" Then Exit Sub
  dbProdutos.Refresh
  If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
  dbProdutos.Recordset.MoveFirst
  dbProdutos.Recordset.Find "codigo=" & .Text
  If dbProdutos.Recordset.EOF = False Then
    cboProduto.Text = dbProdutos.Recordset!Descri
    txtCodigo.Text = dbProdutos.Recordset!Codigo
  End If
End With
End Sub

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
DataGrid1.Visible = False
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

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
DataGrid1.Visible = False
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

