VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc dbVenda 
      Height          =   495
      Left            =   2040
      Top             =   2760
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from venda2"
      Caption         =   "dbVenda"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   125763585
      CurrentDate     =   41944
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Corrige Codigo Venda"
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   125763585
      CurrentDate     =   41974
   End
   Begin MSAdodcLib.Adodc dbFechamento 
      Height          =   495
      Left            =   2040
      Top             =   3360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from fechamento"
      Caption         =   "dbFechamento"
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
   Begin VB.Label lblData 
      Caption         =   "Data Atual"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblAlterados 
      Caption         =   "Alterados: 0"
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Período:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Alterados As Integer

Dim Dia As Date
Dim codFuncionario As String, TotalProduto As Currency
Dim StrTemp As String
Dim Codigo As String, Descri As String, Tipo As String, Valor As Currency
Dim ValorBruto As Currency, Tarifa As Currency, Operacao As Currency
Dim TotalOper As Double, Porcento As Double, Liquido As Currency
Dim DescontoPorcento As Currency
Dim Tanque As Integer, Estoque As Double
Dim Bico As Integer, Encerrante As Double, Encontrou As Boolean, Abertura As Double
Dim Preco As Currency, Qtd As Double, Funcionario As Integer
Dim CodigoConta As String, DesteCaixaQtd As Double, DesteCaixaValor As Currency

Dim DataCaixa As Date, Turno As String


Dim CodigoCliente As Double, Cupom As String, Placa As String
Dim Km As String, Veiculo As String, ValorTotal As Currency
Dim CodigoProduto As Double, valorUnitario As Currency
Dim ValorUnitarioDif As Currency, ValorTotalDif As Currency, LucroDif As Currency
Dim PrecoDif As Boolean, TempValorPagar As Currency
Dim Autorizar As Boolean, Motivo As String, Autorizado As Boolean
Dim Documento As String, DataBordero As Date


Dim db As New ADODB.Connection
Dim dbSql As New ADODB.Connection
Dim dbConfig As New ADODB.Recordset
Dim dbImportacao As New ADODB.Recordset
Dim dbVendedores As New ADODB.Recordset



Dim Caminho As String

On Error Resume Next
With CommonDialog1
  .ShowOpen
  Caminho = .FileName
  If Err.Number <> 0 Then GoTo Sair
End With
On Error GoTo 0

db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Caminho


dbConfig.CursorLocation = adUseClient
dbConfig.Open "select *from config", db, adOpenForwardOnly, adLockReadOnly

dbFechamento.ConnectionString = db.ConnectionString
dbFechamento.RecordSource = "select *from FechamentodeCaixa where datacaixa between #" & txtDataIni.Value & "# and #" & txtDataFim.Value & "# order by datacaixa, horaini"
dbFechamento.Refresh

dbVendedores.CursorLocation = adUseClient
dbVendedores.Open "select *from vendedores", db, adOpenForwardOnly, adLockOptimistic



If dbFechamento.Recordset.RecordCount = 0 Then
  MsgBox "Não existe caixa para esse período!"
  Exit Sub
End If



dbSql.Open "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & "escricentral.dyndns.org,92" 'dbConfig!ftp
dbImportacao.CursorLocation = adUseClient
On Error Resume Next
  
  dbImportacao.Open "select *from caixas where linhaexportada like '002%' and datacaixa between '" & txtDataIni.Value & "' and '" & txtDataFim.Value & "' and codigoposto='" & dbConfig!Porta & "'  order by dataCaixa, turno, linhaexportada", dbSql, adOpenForwardOnly, adLockReadOnly
  
  If Err.Number <> 0 Then
    MsgBox Err.Number & " - " & Err.Description
  End If
  
  On Error GoTo 0
  
  If dbImportacao.RecordCount = 0 Then
    MsgBox "Período não encontrado!"
    GoTo Sair
  End If
  dbImportacao.MoveLast
  dbImportacao.MoveFirst
  
  Do While dbImportacao.EOF = False
    StrTemp = dbImportacao!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
      Case "002"
        'Grava Venda
        DataCaixa = dbImportacao!DataCaixa
        Turno = dbImportacao!Turno
        lblData.Caption = "Data atual: " & DataCaixa
        lblData.Refresh
        
        dbFechamento.Recordset.MoveLast
        dbFechamento.Recordset.MoveFirst
        dbFechamento.Recordset.Find "datacaixa=#" & DataInglesa(DataCaixa) & "#"
        If dbFechamento.Recordset.EOF = True Then
          GoTo naoIncuirProduto
        End If
        If dbFechamento.Recordset!Turno <> Turno Then
          dbFechamento.Recordset.Find "turno='" & Turno & "'"
          If dbFechamento.Recordset.EOF = False Then
            If DataInglesa(DataCaixa) <> dbFechamento.Recordset!DataCaixa Then
              GoTo naoIncuirProduto
            End If
          End If
        End If
        If Trim(Mid(StrTemp, 18, 6)) <> "" Then
          Bico = CInt(Mid(StrTemp, 18, 6))
          GoTo naoIncuirProduto
        Else
          Bico = 0
        End If
        Preco = CCur(Mid(StrTemp, 38, 12))
        StrTemp2 = Mid(StrTemp, 5, 12)
        If IsNumeric(StrTemp2) = False Then
          StrTemp2 = RemoveString(StrTemp2)
        End If
        Codigo = CDbl(StrTemp2)
        Qtd = CDbl(Mid(StrTemp, 25, 12))
        If Trim(Mid(StrTemp, 64)) <> "" Then
          Funcionario = CInt(Mid(StrTemp, 64))
        Else
          Funcionario = 0
        End If
        If Qtd = 0 Then
          GoTo naoIncuirProduto
        End If
        StrTemp2 = Mid(StrTemp, 51, 12)
        If IsNumeric(StrTemp2) = True Then
          TotalProduto = CCur(StrTemp2)
        Else
          TotalProduto = 0
        End If
        
        dbVenda.ConnectionString = db.ConnectionString
        dbVenda.RecordSource = "select *from venda2 where codigofechamento=" & dbFechamento.Recordset!codigofechamento & " and codproduto=" & Codigo & " and quantidade=" & Qtd & " and codigopagamento=" & Funcionario
        dbVenda.Refresh
        
        If dbVenda.Recordset.RecordCount <> 0 Then
          dbVendedores.MoveLast
          dbVendedores.MoveFirst
          dbVendedores.Find "codigo=" & Funcionario
          
          If dbVendedores.EOF = False Then
            dbVenda.Recordset!codigovendedor = dbVendedores!codigovendedor
          End If
          dbVenda.Recordset!codigopagamento = Funcionario
          dbVenda.Recordset.UpdateBatch adAffectAll
          Alterados = Alterados + 1
          lblAlterados.Caption = "Alterados: " & Alterados
        End If
        
        
        
naoIncuirProduto:
    End Select
    dbImportacao.MoveNext
  Loop

Sair:

dbConfig.Close
dbVendedores.Close

MsgBox "Finalizado!"

End Sub

Public Function RemoveString(ByVal TextoOrigem As String, Optional StrARemover As String = "") As String
Dim A As Double
If StrARemover <> "" Then
  A = 0
  A = InStr(1, TextoOrigem, StrARemover)
  Do While A <> 0
    TextoOrigem = Mid(TextoOrigem, 1, A - 1) & Mid(TextoOrigem, A + 1)
    A = 0
    A = InStr(1, TextoOrigem, StrARemover)
  Loop
  RemoveString = TextoOrigem
Else
  RemoveString = ""
  For A = 1 To Len(TextoOrigem)
    If IsNumeric(Mid(TextoOrigem, A, 1)) = True Then
      RemoveString = RemoveString & Mid(TextoOrigem, A, 1)
    End If
  Next A
End If
End Function

Public Function DataInglesa(ByVal Data As String) As String
On Error Resume Next
Data = Format(Month(Data), "00") & "/" & Format(Day(Data), "00") & "/" & Format(Year(Data), "0000")
'If Len(Data) = 10 Then
'    Data = Trim(Mid(Data, 4, 3)) & Trim(Mid(Data, 1, 3)) & Trim(Mid(Data, 7, 4))
'Else
'    Data = Trim(Mid(Data, 4, 3)) & Trim(Mid(Data, 1, 3)) & Trim(Mid(Data, 7, 2))
'End If
DataInglesa = Data
End Function

