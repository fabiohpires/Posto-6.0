VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatFaturamentoClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faturamento de Clientes"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12690
   Icon            =   "frmRelatFaturamentoClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   12690
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   495
      Left            =   9360
      Picture         =   "frmRelatFaturamentoClientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar Sugestão"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc qFaturamento 
      Height          =   375
      Left            =   2280
      Top             =   3960
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select sum(totalcupom) as total from clientesfaturadoperiodo"
      Caption         =   "qFaturamento"
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
   Begin MSAdodcLib.Adodc dbClientes 
      Height          =   375
      Left            =   2280
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from clientes"
      Caption         =   "dbClientes"
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
   Begin MSAdodcLib.Adodc dbNotas 
      Height          =   375
      Left            =   2280
      Top             =   3240
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from clientesnota2 where codigoclientenota=0"
      Caption         =   "dbNotas"
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
   Begin MSAdodcLib.Adodc dbFaturamento 
      Height          =   375
      Left            =   2280
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from clientesfaturadoperiodo"
      Caption         =   "dbFaturamento"
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
      Bindings        =   "frmRelatFaturamentoClientes.frx":0EC4
      Height          =   5775
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   10186
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "CodigoCliente"
         Caption         =   "Cod."
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
         DataField       =   "Nome"
         Caption         =   "Nome"
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
         DataField       =   "UltimoAbastecimento"
         Caption         =   "Ultimo Ab."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LimiteAtual"
         Caption         =   "Limite Atual"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "TotalCupom"
         Caption         =   "Total de Cupons"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "DiasAbastecendo"
         Caption         =   "Tipo"
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
      BeginProperty Column06 
         DataField       =   "Prazo"
         Caption         =   "Prazo"
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
      BeginProperty Column07 
         DataField       =   "MediaDiaria"
         Caption         =   "Media Diária"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "porcento"
         Caption         =   "+ %"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Sugestao"
         Caption         =   "Sugestão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
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
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   450,142
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3855,118
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1319,811
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   434,835
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   524,976
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   569,764
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtPorcento 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5880
      TabIndex        =   5
      Text            =   "20"
      Top             =   120
      Width           =   495
   End
   Begin MSComCtl2.DTPicker txtDataini 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58982401
      CurrentDate     =   37683
   End
   Begin MSComCtl2.DTPicker txtDatafim 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   58982401
      CurrentDate     =   37683
   End
   Begin VB.Label lblTotalCupons 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qFaturamento"
      Height          =   255
      Left            =   10920
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total de cupons:"
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "%"
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Acrécimo para sugestão:"
      Height          =   255
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatFaturamentoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdem As String

Private Sub cmdAplicar_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja aplicar o limite sugerido nos clientes?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With dbFaturamento
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbClientes.Recordset.MoveFirst
      dbClientes.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
      If dbClientes.Recordset.EOF = False Then
        dbClientes.Recordset!Limite = .Recordset!Sugestao
        dbClientes.Recordset.Update
      End If
      .Recordset!limiteatual = .Recordset!Sugestao
      .Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim Db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset
Dim Dias As Double, Sugestao As Currency

If IsNumeric(txtPorcento.Text) = False Then
  MsgBox "Informe um valor correto!"
  txtPorcento.SetFocus
  Exit Sub
End If

Db.Open CaminhoADO
Db.Execute "delete ClientesFaturadoPeriodo.* from ClientesFaturadoPeriodo"
Db.Close

dbFaturamento.Refresh

Dias = DateDiff("d", txtDataini.Value, txtDatafim.Value)

With dbClientes
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    With dbNotas
      .RecordSource = "select codigocliente, sum(valorprevisto) as Total from clientesNota2 where data between #" & DataInglesa(txtDataini.Value) & "# and #" & DataInglesa(txtDatafim.Value) & "# group by codigocliente"
      .Refresh
    End With
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbFaturamento.Recordset.AddNew
      dbFaturamento.Recordset!DataIni = txtDataini.Value
      dbFaturamento.Recordset!DataFim = txtDatafim.Value
      dbFaturamento.Recordset!CodigoCliente = .Recordset!CodigoCliente
      dbFaturamento.Recordset!Nome = .Recordset!Nome
      dbFaturamento.Recordset!prazo = .Recordset!Praso
      dbFaturamento.Recordset!Porcento = txtPorcento.Text
      dbFaturamento.Recordset!limiteatual = .Recordset!Limite
      dbFaturamento.Recordset!UltimoAbastecimento = .Recordset!UltimoAbastecimento
      Select Case .Recordset!Tipo
        Case "Semanal"
          dbFaturamento.Recordset!diasabastecendo = 7
        Case "Quinzenal"
          dbFaturamento.Recordset!diasabastecendo = 15
        Case "Mensalista"
          dbFaturamento.Recordset!diasabastecendo = 30
        Case "Antecipado"
          dbFaturamento.Recordset!diasabastecendo = 0
        Case Else
          dbFaturamento.Recordset!diasabastecendo = 0
      End Select
      dbFaturamento.Recordset!mediadiaria = 0
      dbFaturamento.Recordset!Sugestao = 0
      dbFaturamento.Recordset!totalcupom = 0
      If dbNotas.Recordset.RecordCount <> 0 Then
        dbNotas.Recordset.MoveFirst
        dbNotas.Recordset.Find "codigocliente=" & .Recordset!CodigoCliente
        If dbNotas.Recordset.EOF = False Then
          dbFaturamento.Recordset!mediadiaria = dbNotas.Recordset!Total / Dias
          dbFaturamento.Recordset!totalcupom = dbNotas.Recordset!Total
          
          Sugestao = (dbFaturamento.Recordset!mediadiaria * (dbFaturamento.Recordset!diasabastecendo + dbFaturamento.Recordset!prazo))
          Sugestao = Sugestao + (Sugestao * (CDbl(txtPorcento.Text) / 100))
          dbFaturamento.Recordset!Sugestao = Sugestao
          dbFaturamento.Recordset.Update
        End If
      End If
      .Recordset.MoveNext
    Loop
  End If
  
End With

With dbFaturamento
  .Refresh
  .Refresh
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With qFaturamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(totalcupom) as total from clientesfaturadoperiodo where dataini>=#" & DataInglesa(txtDataini.Value) & "# and datafim<=#" & DataInglesa(txtDatafim.Value) & "#"
  .Refresh
  .Refresh
End With

End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String

On Error GoTo naoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = "Período: " & txtDataini.Value & " a " & txtDatafim.Value & "  -  Taxa de Acrécimo: " & txtPorcento.Text & " %"
StrTemp = StrTemp & Chr(vbKeyReturn) & "Impresso em: " & Format(Now, "long date") & " - " & Format(Now, "short time")

Printer.Orientation = vbPRORLandscape

ImprimeADOGrid DataGrid1, Printer, dbFaturamento, 4, True, , , , , NomePosto, "Faturamento de Clientes", StrTemp

Printer.EndDoc

naoImprime:

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField Then
  StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  StrOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField
End If

With dbFaturamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from clientesfaturadoperiodo" & StrOrdem
  .Refresh
End With

End Sub

Private Sub Form_DblClick()
dbFaturamento.Refresh
End Sub

Private Sub Form_Load()
Dim Db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

txtDataini.Value = DateAdd("m", -1, Date)
txtDatafim.Value = Date

Db.Open CaminhoADO

StrOrdem = " order by Nome"

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesFaturadoPeriodo", Db
If Err.Number <> 0 Then
  dbTemp.Close
  On Error GoTo 0
  On Error Resume Next
  Db.Execute "create table ClientesFaturadoPeriodo (DataIni datetime, DataFim datetime, CodigoCliente double, Nome Text(100), TotalCupom currency, DiasAbastecendo integer, Prazo integer, MediaDiaria currency, porcento double, Sugestao currency)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao criar a tabela 'ClientesFaturadoPeriodo'!"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesFaturadoPeriodo order by UltimoAbastecimento", Db
If Err.Number <> 0 Then
  dbTemp.Close
  On Error GoTo 0
  On Error Resume Next
  Db.Execute "alter table ClientesFaturadoPeriodo add column UltimoAbastecimento datetime"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesFaturadoPeriodo->UltimoAbastecimento'!"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ClientesFaturadoPeriodo order by LimiteAtual", Db
If Err.Number <> 0 Then
  dbTemp.Close
  On Error GoTo 0
  On Error Resume Next
  Db.Execute "alter table ClientesFaturadoPeriodo add column LimiteAtual Currency"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description & " ao alterar a tabela 'ClientesFaturadoPeriodo->LimiteAtual'!"
  End If
End If
dbTemp.Close

On Error GoTo 0

Db.Close

With dbFaturamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from clientesfaturadoperiodo" & StrOrdem
  .Refresh
End With
With dbNotas
  .ConnectionString = CaminhoADO
End With
With dbClientes
  .ConnectionString = CaminhoADO
End With
With qFaturamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select sum(totalcupom) as total from clientesfaturadoperiodo where dataini>=#" & DataInglesa(txtDataini.Value) & "# and datafim<=#" & DataInglesa(txtDatafim.Value) & "#"
  .Refresh
End With


End Sub
