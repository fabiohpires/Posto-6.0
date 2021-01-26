VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "Msdatgrd.ocx"
Begin VB.Form frmcadProdutosConfere 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conferência de Estoque"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10740
   Icon            =   "frmProdutosConfere.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10740
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSemCombustivel 
      Caption         =   "Sem Combustíveis"
      Height          =   255
      Left            =   5880
      TabIndex        =   11
      Top             =   120
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton cmdErros 
      Caption         =   "Erros"
      Height          =   375
      Left            =   8520
      TabIndex        =   10
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   240
      Picture         =   "frmProdutosConfere.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Imprimir"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CheckBox chkDiferenca 
      Caption         =   "Somente com Diferença"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Arquivo de Estoque"
      FileName        =   "Estoque.txt"
      Filter          =   "Arquivo Texto *.txt|*.txt"
   End
   Begin VB.CommandButton cmdImporta 
      Caption         =   "Importação"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox txtIdPosto 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Text            =   "01"
      Top             =   120
      Width           =   375
   End
   Begin MSAdodcLib.Adodc dbProdutosCodigos 
      Height          =   330
      Left            =   3600
      Top             =   4440
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "select *from ProdutosCodigos"
      Caption         =   "dbProdutosCodigos"
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
   Begin MSAdodcLib.Adodc qTotal 
      Height          =   330
      Left            =   3600
      Top             =   4800
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "select *from ProdutosCodigos"
      Caption         =   "qTotal"
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
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   3600
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      RecordSource    =   "select *from Produtos order by descri"
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProdutosConfere.frx":0EC4
      Height          =   5775
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   10186
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Código"
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
         DataField       =   "Descri"
         Caption         =   "Descrição"
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
         DataField       =   "Estoque"
         Caption         =   "Estoque"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "EstoqueFisico"
         Caption         =   "Estoque Físico"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Diferenca"
         Caption         =   "Diferença"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "PrecoCompra"
         Caption         =   "$ Compra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PrecoVenda"
         Caption         =   "$ Venda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3060,284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1230,236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1019,906
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltotalCusto 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   6480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Custo Total:"
      Height          =   255
      Left            =   7680
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lblDifTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6480
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Diferença Total:"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Id do Posto:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmcadProdutosConfere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String, strFiltro As String

Private Sub Atualiza()
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos" & strFiltro & StrOrdem
  .Refresh
End With
dbProdutosCodigos.Refresh

With qTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(diferenca) as totalDiferenca, sum(diferenca*precocompra) as totalCusto from produtos" & strFiltro
  .Refresh
  If IsNull(.Recordset!totaldiferenca) = False Then
    lblDifTotal.Caption = .Recordset!totaldiferenca
  Else
    lblDifTotal.Caption = "0"
  End If
  If IsNull(.Recordset!TotalCusto) = False Then
    lbltotalCusto.Caption = Format(.Recordset!TotalCusto, "Currency")
  Else
    lbltotalCusto.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub chkDiferenca_Click()
If chkDiferenca.Value = vbChecked Then
  strFiltro = " where diferenca<>0"
Else
  strFiltro = ""
End If
If chkSemCombustivel.Value = vbChecked Then
  If strFiltro = "" Then
    strFiltro = " where combustivel=0"
  Else
    strFiltro = strFiltro & " and combustivel=0"
  End If
End If
Atualiza

End Sub

Private Sub chkSemCombustivel_Click()
Call chkDiferenca_Click
End Sub

Private Sub cmdErros_Click()
Shell "notepad C:\erros.txt", vbNormalFocus
End Sub

Private Sub cmdImporta_Click()


Dim Ws As Workspace, Db As Database
Dim StrArquivo As String, intArquivo As Integer
Dim strTemp As String
Dim Codigo As Double, Quantidade As Double, CodBar As String, Loja As String
Dim Posto As String, CodigoProduto As Double

Set Ws = DBEngine.Workspaces(0)
Set Db = Ws.OpenDatabase(Caminho, , , Conectar)
Db.Execute "update produtos set estoquefisico=0"
Db.Execute "update produtos set produtos.diferenca=-estoque"

If txtIdPosto.Text = "" Then
  MsgBox "Informe o Id do Posto!"
  txtIdPosto.SetFocus
  Exit Sub
End If
Posto = txtIdPosto.Text
On Error GoTo TrataErro
With CommonDialog1
  .ShowOpen
  StrArquivo = .FileName
  If Dir(StrArquivo) = "" Then
    MsgBox "Arquivo não encontrado!"
    Exit Sub
  End If
End With
On Error GoTo 0
intArquivo = FreeFile()
Open StrArquivo For Input As intArquivo
Do While EOF(intArquivo) = False
  Line Input #intArquivo, strTemp
  If strTemp <> "" Then
    Loja = Mid(strTemp, 1, 2)
    CodBar = Trim(Mid(strTemp, 11, 25))
    If IsNumeric(Trim(Mid(strTemp, 37, 6))) = True Then
      Quantidade = CDbl(Trim(Mid(strTemp, 37, 6)))
    Else
      a = FreeFile()
      Open "c:\Erros.txt" For Append As #a
      Print #a, strTemp & " quantidade inválida"
      Close #a
      CodigoProduto = 0
      Quantidade = 0
    End If
    If IsNumeric(Trim(Mid(strTemp, 4, 6))) = True Then
      Codigo = CDbl(Trim(Mid(strTemp, 4, 6)))
    Else
      Codigo = 0
    End If
    With dbProdutos
      If .Recordset.RecordCount <> 0 Then
        .Recordset.Find "codigo=" & Codigo
        If .Recordset.EOF = False Then
          CodigoProduto = .Recordset!CodigoProduto
        Else
          a = FreeFile()
          Open "c:\Erros.txt" For Append As #a
          Print #a, strTemp & " não encontrado"
          Close #a
          CodigoProduto = 0
        End If
      End If
    End With
    If Loja = Posto Then
      With dbProdutosCodigos
        .Refresh
        If CodigoProduto <> 0 Then
          If .Recordset.RecordCount = 0 Then
            .Recordset.AddNew
            .Recordset!CodigoProduto = Codigo
            .Recordset!codigobarra = CodBar
            .Recordset!codigosistema = CodigoProduto
            .Recordset.Update
          Else
            .Recordset.MoveFirst
            .Recordset.Find "codigobarra='" & CodBar & "'"
            If .Recordset.EOF = True Then
              .Recordset.AddNew
              .Recordset!CodigoProduto = Codigo
              .Recordset!codigobarra = CodBar
              .Recordset!codigosistema = CodigoProduto
              .Recordset.Update
            Else
              If .Recordset!CodigoProduto <> CodigoProduto Then
                .Recordset!CodigoProduto = Codigo
                .Recordset!codigobarra = CodBar
                .Recordset!codigosistema = CodigoProduto
                .Recordset.Update
              End If
            End If
          End If
        Else
          If .Recordset.RecordCount = 0 Then
            MsgBox "Não existe codigo de barras cadastrado!"
            Exit Sub
          End If
          .Recordset.MoveFirst
          .Recordset.Find "codigobarra='" & CodBar & "'"
          If .Recordset.EOF = False Then
            Codigo = .Recordset!CodigoProduto
            If IsNull(.Recordset!codigosistema) = False Then
              CodigoProduto = .Recordset!codigosistema
            Else
              CodigoProduto = 0
            End If
          End If
        End If
      End With
      With dbProdutos
        .Recordset.MoveFirst
        .Recordset.Find "codigo=" & Codigo
        If .Recordset.EOF = False Then
          .Recordset!estoquefisico = .Recordset!estoquefisico + Quantidade
          .Recordset!Diferenca = .Recordset!estoquefisico - .Recordset!Estoque
          .Recordset.Update
        End If
      End With
    End If
  End If
Loop
TrataErro:
Close #intArquivo
Atualiza
End Sub

Private Sub cmdImprime_Click()
On Error GoTo naoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

ImprimeADOGrid DataGrid1, Printer, dbProdutos, 4, , , , , , "Diferença de Estoque", NomePosto, Chr(13) & "Impresso em: " & Format(Date, "longdate")

Printer.Print
Printer.Print "Diferença Total: " & lblDifTotal.Caption
Printer.Print "Custo Total: " & lbltotalCusto.Caption
Printer.EndDoc

naoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  StrOrdem = " order by " & DBGrid1.Columns(ColIndex).DataField
End If

If chkDiferenca.Value = vbChecked Then
  strFiltro = " where diferenca<>0"
Else
  strFiltro = ""
End If

Atualiza

End Sub

Private Sub Form_Load()

StrOrdem = " order by Codigo"
strFiltro = " where combustivel=0"
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos" & StrOrdem
  .Refresh
End With
With dbProdutosCodigos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With qTotal
  .ConnectionString = CaminhoADO
  .Refresh
End With
Atualiza
End Sub
