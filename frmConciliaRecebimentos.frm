VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConciliaRecebimentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliação de Recebimentos"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "frmConciliaRecebimentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPesquisa 
      Caption         =   "Pesquisar"
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtValorRec 
      Height          =   285
      Left            =   3120
      TabIndex        =   15
      Top             =   360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dbFormaDePg 
      Height          =   330
      Left            =   3360
      Top             =   3600
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from formadepagamento order by descri"
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
   Begin MSDataListLib.DataCombo cboTipoRec 
      Bindings        =   "frmConciliaRecebimentos.frx":0442
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc QTemp 
      Height          =   330
      Left            =   3360
      Top             =   3240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from previsaorecebimentosSoma"
      Caption         =   "QTemp"
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
   Begin MSAdodcLib.Adodc dbPendenciasSoma 
      Height          =   330
      Left            =   3360
      Top             =   2880
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from previsaorecebimentosSoma"
      Caption         =   "dbPendenciasSoma"
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
      Left            =   6600
      TabIndex        =   8
      Top             =   240
      Width           =   615
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
      Left            =   5880
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   6000
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   24444929
      CurrentDate     =   37648
   End
   Begin MSAdodcLib.Adodc dbPendencias 
      Height          =   330
      Left            =   3360
      Top             =   2520
      Visible         =   0   'False
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from previsaorecebimentos"
      Caption         =   "dbPendencias"
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
      Bindings        =   "frmConciliaRecebimentos.frx":045C
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5953
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "DataEntrada"
         Caption         =   "Lançada"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DataPrevista"
         Caption         =   "Prevista"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Descri"
         Caption         =   "Forma de Pg."
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
         DataField       =   "ValorBruto"
         Caption         =   "Bruto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "ValorLiquidoPrevisto"
         Caption         =   "Liq. Prev."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ ""#.##0,00"
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
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1110,047
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmConciliaRecebimentos.frx":0477
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2566
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Descri"
         Caption         =   "Descri"
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
         DataField       =   "Bruto"
         Caption         =   "Bruto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LiquidoPrev"
         Caption         =   "LiquidoPrev"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3240
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1110,047
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1094,74
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Valor:"
      Height          =   255
      Left            =   3120
      TabIndex        =   14
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo Recebido:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   6120
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5880
      Width           =   405
   End
   Begin VB.Label Label2 
      Caption         =   "Valor Rec.:"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   1680
      TabIndex        =   0
      Top             =   5880
      Width           =   390
   End
End
Attribute VB_Name = "frmConciliaRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoSoma As String

Private Sub cmdConfirmar_Click()
Dim Dia As Date
If dbPendencias.Recordset.RecordCount = 0 Then Exit Sub
If dbPendencias.Recordset.EOF = True Then
  MsgBox "Selecione um registro primeiro!"
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Informe um valor válido!"
  txtValor.SetFocus
  Exit Sub
End If
If dbPendenciasSoma.Recordset.RecordCount = 0 Then
  MsgBox "Escolha pelo menos um lançamento para ser confirmado!"
  Exit Sub
End If
Dia = txtData.Value
With QTemp
  .RecordSource = "select codigosoma, sum(liquidoprev) as liquido from previsaoRecebimentosSoma where codigosoma='" & CodigoSoma & "' group by codigosoma"
  .Refresh
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Soma = .Recordset!Liquido
  Else
    Soma = 0
  End If
End With

With frmConciliacao
  With .dbConcilia
    .Recordset.AddNew
    .Recordset!codigoconta = frmConciliacao.dbContas.Recordset!codigoconta
    .Recordset!Data = txtData.Value
    .Recordset!tipo = "Recebimento"
    .Recordset!Codigo = dbPendencias.Recordset!codigoprevisaorecebe
    .Recordset!Descri = dbPendencias.Recordset!Descri
    Descri = dbPendencias.Recordset!Descri
    .Recordset!NrDocumento = "111111111"
    .Recordset!Valor = CCur(txtValor.Text)
    .Recordset.Update
    .Refresh
    .Refresh
  End With
  With .dbContas
    codigoconta = .Recordset!codigoconta
    .Refresh
    .Recordset.Find "codigoconta=" & codigoconta
    If .Recordset.EOF = False Then
      .Recordset!Saldo = .Recordset!Saldo + CCur(txtValor.Text)
      .Recordset.Update
    End If
  End With
  TempValor = CCur(txtValor.Text)
  
  With dbPendenciasSoma
    .Refresh
    Do While .Recordset.EOF = False
      With dbPendencias
        .Refresh
        .Recordset.Find "codigoprevisaorecebe=" & dbPendenciasSoma.Recordset!codigopendencia
        If .Recordset.EOF = False Then
          .Recordset!Confirmado = True
          .Recordset!datarecebida = txtData.Value
          .Recordset!dataconfirmada = Now
          .Recordset!valorconfirmado = TempValor
          .Recordset!difrecebido = TempValor - Soma
          .Recordset!fechadiferenca = False
          .Recordset.Update
        End If
        TempValor = 0
        Soma = 0
        .Refresh
        .Refresh
      End With
      .Recordset.MoveNext
    Loop
    .Recordset.MoveLast
    .Recordset!liquidorecebido = CCur(txtValor.Text)
    .Recordset.Update
    dbPendencias.Refresh
    dbPendencias.Recordset.Find "codigoprevisaorecebe=" & dbPendenciasSoma.Recordset!codigopendencia
    If dbPendencias.Recordset.EOF = False Then
      dbPendencias.Recordset!ValorRecebido = CCur(txtValor.Text)
      dbPendencias.Recordset.Update
    End If
  End With
  With .dbMovimentacao
    .Recordset.AddNew
    .Recordset!Data = Now
    .Recordset!tipo = "Conciliação"
    .Recordset!codigoconta = frmConciliacao.dbContas.Recordset!codigoconta
    .Recordset!conta = frmConciliacao.dbContas.Recordset!Descri
    .Recordset!Descri = Descri & " "
    .Recordset!Valor = CCur(txtValor.Text)
    .Recordset!Saldo = frmConciliacao.dbContas.Recordset!Saldo
    .Recordset.Update
    .Refresh
    .Refresh
  End With
  
  .TiraSaldo Dia
End With

CodigoSoma = Trim(Str(CDbl(Now)))
With dbPendenciasSoma
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from PrevisaoRecebimentosSoma where codigosoma='" & CodigoSoma & "'"
  .Refresh
End With

With QTemp
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select codigosoma, sum(liquidoprev) as liquido from previsaoRecebimentosSoma where codigosoma='" & CodigoSoma & "' group by codigosoma"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    lblTotal.Caption = Format(.Recordset!Liquido, "currency")
  Else
    lblTotal.Caption = Format(0, "currency")
  End If
End With

DataGrid1.SetFocus
If dbPendencias.Recordset.RecordCount <> 0 Then
  If dbPendencias.Recordset.EOF = True Then
    dbPendencias.Recordset.MoveFirst
  End If
End If
End Sub

Private Sub cmdPesquisa_Click()
With dbFormaDePg
  .Refresh
  If cboTipoRec.Text = "" Then
    MsgBox "Escolha um tipo de Recebimento!"
    cboTipoRec.SetFocus
    Exit Sub
  End If
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe forma de pagamento cadastrado!"
    Exit Sub
  End If
  .Recordset.Find "descri='" & cboTipoRec.Text & "'"
  If .Recordset.EOF = True Then
    MsgBox "Forma de pagamento não encontrada!"
    cboTipoRec.SetFocus
  End If
End With
If IsNumeric(txtValorRec.Text) = False Then
  MsgBox "Informe um valor válido"
  txtValorRec.SetFocus
  Exit Sub
End If

With dbPendencias
  txtValor.Text = Format("0", "Currency")
  
  .Refresh
  .Recordset.Filter = "grupo=" & dbFormaDePg.Recordset!grupo
  .Recordset.Sort = "dataentrada"
  Do While .Recordset.EOF = False
    TempValor = CCur(txtValor.Text) - CCur(txtValorRec.Text)
    If TempValor < 10 Then
      Call cmdSomar_Click
    End If
    TempValor = CCur(txtValor.Text) - CCur(txtValorRec.Text)
    If TempValor > 10 Then
      Call cmdSomar_Click
      Call cmdSomar_Click
      Call cmdSubtrair_Click
      Exit Do
    End If
    .Recordset.MoveNext
  Loop
  .Refresh
End With


End Sub

Private Sub cmdSair_Click()
With dbPendenciasSoma
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.RecordCount <> 0
      .Recordset.Delete
      .Refresh
      .Refresh
    Loop
  End If
End With
Unload Me
End Sub

Private Sub cmdSomar_Click()
If dbPendencias.Recordset.RecordCount = 0 Then Exit Sub
If dbPendencias.Recordset.EOF = True Then Exit Sub
With dbPendenciasSoma
  .Refresh
  .Recordset.Find "Codigopendencia=" & dbPendencias.Recordset!codigoprevisaorecebe
  If .Recordset.EOF = False Then
    Exit Sub
  End If
  .Recordset.AddNew
  .Recordset!CodigoSoma = CodigoSoma
  .Recordset!codigopendencia = dbPendencias.Recordset!codigoprevisaorecebe
  .Recordset!Descri = dbPendencias.Recordset!Descri
  .Recordset!Data = dbPendencias.Recordset!DataPrevista
  .Recordset!Bruto = dbPendencias.Recordset!ValorBruto
  .Recordset!liquidoprev = dbPendencias.Recordset!valorliquidoprevisto
  .Recordset.Update
  .Refresh
  .Refresh
End With
With QTemp
  .RecordSource = "select codigosoma, sum(liquidoprev) as liquido from previsaoRecebimentosSoma where codigosoma='" & CodigoSoma & "' group by codigosoma"
  .Refresh
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    lblTotal.Caption = Format(.Recordset!Liquido, "currency")
  Else
    lblTotal.Caption = Format(0, "currency")
  End If
End With

txtValor.Text = Format(lblTotal.Caption, "Currency")
txtData.Value = dbPendencias.Recordset!DataPrevista
DataGrid1.SetFocus

End Sub

Private Sub cmdSubtrair_Click()
With dbPendenciasSoma
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  .Recordset.Delete
  .Refresh
  .Refresh
  .Refresh
End With
With QTemp
  .RecordSource = "select codigosoma, sum(liquidoprev) as liquido from previsaoRecebimentosSoma where codigosoma='" & CodigoSoma & "' group by codigosoma"
  .Refresh
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    lblTotal.Caption = Format(.Recordset!Liquido, "currency")
    
  Else
    lblTotal.Caption = Format(0, "currency")
  End If
End With
txtValor.Text = lblTotal.Caption
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If dbPendencias.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  dbPendencias.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  dbPendencias.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
  Case Asc("-")
    Call cmdSubtrair_Click
  Case Asc("+")
    Call cmdSomar_Click
End Select

End Sub

Private Sub Form_Load()
With dbPendencias
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbFormaDePg
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With

txtData.Value = Date
CodigoSoma = Trim(Str(CDbl(Now)))
With dbPendenciasSoma
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from PrevisaoRecebimentosSoma where codigosoma='" & CodigoSoma & "'"
  .Refresh
End With

With QTemp
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select codigosoma, sum(liquidoprev) as liquido from previsaoRecebimentosSoma where codigosoma='" & CodigoSoma & "' group by codigosoma"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    lblTotal.Caption = Format(.Recordset!Liquido, "currency")
  Else
    lblTotal.Caption = Format(0, "currency")
  End If
End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValorRec_GotFocus()
With txtValorRec
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValorRec_LostFocus()
With txtValorRec
  If .Text = "" Then Exit Sub
  If IsNumeric(.Text) = False Then Exit Sub
  .Text = Format(.Text, "Currency")
End With
End Sub
