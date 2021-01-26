VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFaturaClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Faturamento de Clientes"
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmFaturaClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   20
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   5400
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbClientesCobrancaTemp 
      Height          =   375
      Left            =   960
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from ClientesCobrancaTemp"
      Caption         =   "dbClientesCobrancaTemp"
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
   Begin MSAdodcLib.Adodc dbClientesNotaTemp 
      Height          =   375
      Left            =   960
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from ClientesNotaTemp"
      Caption         =   "dbClientesNotaTemp"
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
   Begin MSDataListLib.DataCombo cboCliente 
      Bindings        =   "frmFaturaClientes.frx":0442
      Height          =   315
      Left            =   1320
      TabIndex        =   4
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc DbClientes2 
      Height          =   375
      Left            =   960
      Top             =   3240
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from Clientes order by nome"
      Caption         =   "DbClientes2"
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
   Begin MSAdodcLib.Adodc DbClientes 
      Height          =   375
      Left            =   960
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from Clientes order by nome"
      Caption         =   "DbClientes"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmFaturaClientes.frx":045B
      Height          =   2895
      Left            =   120
      TabIndex        =   17
      Top             =   4680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   5106
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
         DataField       =   "Cupom"
         Caption         =   "Cupom"
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
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "Valor"
         Caption         =   "Valor"
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
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3075,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1184,882
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Cliente a Faturar"
      Height          =   615
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   3975
      Begin VB.CheckBox OptMensal 
         Caption         =   "Mensalista"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox optSemanal 
         Caption         =   "Semanal"
         Height          =   255
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox optQuinzenal 
         Caption         =   "Quinzenal"
         Height          =   255
         Left            =   1440
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtCodigo2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin MSComCtl2.DTPicker txtFechamento 
      Height          =   315
      Left            =   5880
      TabIndex        =   11
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   39643
   End
   Begin MSDataListLib.DataCombo cboCliente2 
      Bindings        =   "frmFaturaClientes.frx":047C
      Height          =   315
      Left            =   1320
      TabIndex        =   9
      Top             =   960
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmFaturaClientes.frx":0496
      Height          =   2415
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4260
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
         DataField       =   "CodigoCliente"
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
         DataField       =   "Vencimento"
         Caption         =   "Vencimento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ValorPrevisto"
         Caption         =   "Valor"
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
            ColumnWidth     =   629,858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4454,929
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1065,26
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1230,236
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Fechamento:"
      Height          =   195
      Left            =   5880
      TabIndex        =   10
      Top             =   720
      Width           =   930
   End
   Begin VB.Label Label4 
      Caption         =   "Final:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Código:"
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   8
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label2 
      Caption         =   "Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label11 
      Caption         =   "Código:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmFaturaClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrOrdemCobranca As String, StrOrdemNotas As String

Private Function FechaClienteNota(ByVal CodigoCliente As Double, ByVal DataFechamento As Date, ByVal Praso As Double) As Boolean
Dim db As New ADODB.Connection, dbSoma As New ADODB.Recordset
Dim dbCobranca As New ADODB.Recordset
Dim ValorTotal As Currency, codigoSoma As String
Dim Resposta As Integer, Vencimento As Date
Dim Confirmado As Double, CodigoCobranca As Double

FechaClienteNota = False

db.Open CaminhoADO
dbSoma.Open "select sum(valorprevisto) as total from clientesNota2 where codigocliente=" & CodigoCliente & " and confirmado=0 and data<=#" & DataInglesa(Trim(Str(DataFechamento))) & "#", db, adOpenKeyset
If IsNull(dbSoma!Total) = False Then
  ValorTotal = dbSoma!Total
End If
dbSoma.Close

If ValorTotal = 0 Then Exit Function

codigoSoma = Str(CDbl(Now))
Vencimento = DateAdd("d", Praso, DataFechamento)

db.Execute "update clientesnota2 set codigosoma='" & codigoSoma & "' where codigocliente=" & CodigoCliente & " and data<=#" & DataInglesa(DataFechamento) & "# and confirmado=0"
db.Execute "update clientesnota2 set confirmado=-1 where codigocliente=" & CodigoCliente & " and data<=#" & DataInglesa(DataFechamento) & "# and confirmado=0"

dbSoma.Open "Select *from clientes where codigocliente=" & CodigoCliente, db, adOpenKeyset, adLockOptimistic
dbCobranca.Open "select *from clientescobranca order by codigocobranca desc", db, adOpenKeyset, adLockOptimistic

dbCobranca.AddNew
dbCobranca!datasoma = Now
dbCobranca!DataFechamento = Vencimento
dbCobranca!codigoSoma = codigoSoma
dbCobranca!CodigoCliente = CodigoCliente
dbCobranca!Cliente = dbSoma!Nome
dbCobranca!Valor = ValorTotal
dbCobranca!Origem = "Fiado"
dbCobranca!Descri = "Fechamento até " & Format(DataFechamento, "Short date")
dbCobranca.Update

dbCobranca.Close

dbSoma!TotalBoleto = dbSoma!TotalBoleto + ValorTotal
dbSoma!TotalNotas = dbSoma!TotalNotas - ValorTotal
dbSoma!Saldo = dbSoma!Limite - dbSoma!TotalNotas - dbSoma!TotalBoleto
dbSoma.Update

dbSoma.Close

db.Close

FechaClienteNota = True
End Function

Private Sub cboCliente_Change()
cboCliente2.Text = cboCliente.Text
End Sub

Private Sub cboCliente_GotFocus()
With cboCliente
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub cboCliente_LostFocus()
With dbClientes
  .Refresh
  .Recordset.Find "nome='" & cboCliente.Text & "'"
  If .Recordset.EOF = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cboCliente2_GotFocus()
With cboCliente2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub cboCliente2_LostFocus()
With dbClientes2
  .Refresh
  .Recordset.Find "nome='" & cboCliente2.Text & "'"
  If .Recordset.EOF = False Then
    cboCliente2.Text = .Recordset!Nome
    txtCodigo2.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cmdConfirmar_Click()
Dim Resposta As Integer
Resposta = MsgBox("Deseja confirmar apenas o cliente selecionado?" & Chr(13) & "Sim - somente selecionado" & Chr(13) & "Não - todos clientes listados" & Chr(13) & "Cancela - sai sem faturar", vbYesNoCancel + vbDefaultButton3)
Select Case Resposta
  Case vbYes
    With dbClientesCobrancaTemp
      If .Recordset.EOF = True Or .Recordset.BOF = True Then
        MsgBox "É preciso selecionar pelo menos um cliente!"
        Exit Sub
      End If
      If FechaClienteNota(.Recordset!CodigoCliente, .Recordset!DataFechamento, .Recordset!Praso) = True Then
        .Recordset!Confirmado = True
        .Recordset.Update
        .Refresh
        .Refresh
      End If
    End With
  Case vbNo
    With dbClientesCobrancaTemp
      If .Recordset.EOF = True Or .Recordset.BOF = True Then
        MsgBox "É preciso selecionar pelo menos um cliente!"
        Exit Sub
      End If
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        If FechaClienteNota(.Recordset!CodigoCliente, .Recordset!DataFechamento, .Recordset!Praso) = True Then
          .Recordset!Confirmado = True
          .Recordset.Update
          .Refresh
          .Refresh
        End If
      Loop
    End With
  Case vbCancel
    Exit Sub
End Select

End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String, Vencimento As Date, CodigoTemp As Double, Total As Currency
Dim db As New ADODB.Connection, strPlanoDeConta As String
Dim dbClientesTemp As New ADODB.Recordset
Dim dbNotasTemp As New ADODB.Recordset

If DateDiff("d", Date, txtFechamento.Value) > 30 Then
  MsgBox "Data muito futura!"
  Exit Sub
End If
If DateDiff("d", Date, txtFechamento.Value) < -30 Then
  MsgBox "Data muito antiga!"
  Exit Sub
End If


'If OptMensal.Value = vbChecked Or optQuinzenal.Value = vbChecked Or optSemanal.Value = vbChecked Then
'  If OptMensal.Value = vbChecked Then
'    If strTemp <> "" Then
'      strTemp = strTemp & ",Mensalista"
'    Else
'      strTemp = "Mensalista"
'    End If
'  End If
'  If optQuinzenal.Value = vbChecked Then
'    If strTemp <> "" Then
'      strTemp = strTemp & ",Quinzenal"
'    Else
'      strTemp = "Quinzenal"
'    End If
'  End If
'  If optSemanal.Value = vbChecked Then
'    If strTemp <> "" Then
'      strTemp = strTemp & ",Semanal"
'    Else
'      strTemp = "Semanal"
'    End If
'  End If
'  strTemp = "select *from clientes where tipo in (" & strTemp & ")"
'End If

Select Case Usuarios.Grupo.ClientesPlanos
  Case "0"
    strPlanoDeConta = ""
  Case ""
    strPlanoDeConta = ""
  Case Else
    strPlanoDeConta = "'" & Usuarios.Grupo.ClientesPlanos & "'"
    strPlanoDeConta = Replace(strPlanoDeConta, ",", "','")
    strPlanoDeConta = " planodeconta in (" & strPlanoDeConta & ")"
End Select


If dbClientes.Recordset.EOF = False And dbClientes.Recordset.BOF = False Then
  If IsNumeric(txtCodigo.Text) = True And IsNumeric(txtCodigo2.Text) = True Then
    If StrTemp = "" Then
      StrTemp = "select *from clientes where "
    Else
      StrTemp = StrTemp & " and "
    End If
    StrTemp = StrTemp & "codigocliente between " & txtCodigo.Text & " and " & txtCodigo2.Text
  End If
End If

If StrTemp = "" Then
  If strPlanoDeConta = "" Then
    StrTemp = "select *from clientes"
  Else
    StrTemp = "select *from clientes where " & strPlanoDeConta
  End If
Else
  StrTemp = StrTemp & " and " & strPlanoDeConta
End If

db.Open CaminhoADO
db.Execute "delete from clientesnotatemp where confirmado=0"
db.Execute "delete from clientescobrancatemp where confirmado=0"
dbClientesTemp.Open StrTemp, db, adOpenKeyset
dbNotasTemp.Open "select clientesnota2.CodigoClienteNota, " & _
                        "clientesnota2.CodigoCliente, " & _
                        "clientesnota2.data, " & _
                        "clientesnota2.cupom, " & _
                        "clientesnota2.valorprevisto, " & _
                        "produtos.codigoproduto, " & _
                        "produtos.codigo, " & _
                        "produtos.descri from clientesnota2, produtos where " & _
                        "clientesnota2.codigoproduto=produtos.codigo and " & _
                        "confirmado=0 and " & _
                        "data<=#" & DataInglesa(txtFechamento.Value) & "#", db, adOpenKeyset

If dbClientesTemp.BOF = False And dbClientesTemp.EOF = False Then
  dbClientesTemp.MoveLast
  dbClientesTemp.MoveFirst
  Do While dbClientesTemp.EOF = False
    Select Case dbClientesTemp!Tipo
      Case "Mensalista"
        If OptMensal.Value = vbUnchecked Then GoTo Procimo
      Case "Quinzenal"
        If optQuinzenal.Value = vbUnchecked Then GoTo Procimo
      Case "Semanal"
        If optSemanal.Value = vbUnchecked Then GoTo Procimo
    End Select
    dbNotasTemp.Filter = "codigocliente=" & dbClientesTemp!CodigoCliente
    If dbNotasTemp.EOF = False And dbNotasTemp.BOF = False Then
      With dbClientesCobrancaTemp
        .Refresh
        If .Recordset.EOF = False And .Recordset.BOF = False Then
          .Recordset.MoveFirst
          .Recordset.Find "codigocliente=" & dbClientesTemp!CodigoCliente
          If .Recordset.EOF = True Then
            .Recordset.AddNew
          End If
        Else
          .Recordset.AddNew
        End If
        Dias = dbClientesTemp!Praso
        Vencimento = DateAdd("d", Dias, txtFechamento.Value)
        .Recordset!CodigoCliente = dbClientesTemp!CodigoCliente
        .Recordset!Nome = dbClientesTemp!Nome
        .Recordset!Tipo = dbClientesTemp!Tipo
        .Recordset!Vencimento = Vencimento
        .Recordset!ValorPrevisto = 0
        .Recordset!Confirmado = 0
        .Recordset!DataFechamento = txtFechamento.Value
        .Recordset!Praso = dbClientesTemp!Praso
        .Recordset.Update
        .RecordSource = "Select *from ClientesCobrancaTemp where confirmado=0" & StrOrdemCobranca
        .Refresh
        .Recordset.Find "codigocliente=" & dbClientesTemp!CodigoCliente
        If .Recordset.EOF = True Then GoTo Procimo
        CodigoTemp = .Recordset!CodigoTemp
      End With
      Total = 0
      dbNotasTemp.MoveLast
      dbNotasTemp.MoveFirst
      Do While dbNotasTemp.EOF = False
        With dbClientesNotaTemp
          .Recordset.AddNew
          .Recordset!codigoclientescobra = CodigoTemp
          .Recordset!codigoclientesnota2 = dbNotasTemp!CodigoClienteNota
          .Recordset!CodigoCliente = dbClientesTemp!CodigoCliente
          .Recordset!Nome = dbClientesTemp!Nome
          .Recordset!Cupom = dbNotasTemp!Cupom
          .Recordset!Data = dbNotasTemp!Data
          .Recordset!Valor = dbNotasTemp!ValorPrevisto
          .Recordset!Codigo = dbNotasTemp("Codigo")
          .Recordset!CodigoProduto = dbNotasTemp("codigoproduto")
          .Recordset!Descri = dbNotasTemp!Descri
          .Recordset!Confirmado = 0
          .Recordset.Update
          Total = Total + dbNotasTemp!ValorPrevisto
        End With
        
        dbNotasTemp.MoveNext
      Loop
      dbClientesCobrancaTemp.Recordset.Update "valorprevisto", Total
    End If
Procimo:
    
    dbClientesTemp.MoveNext
  Loop
End If

dbClientesTemp.Close
dbNotasTemp.Close
db.Close
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid2_HeadClick(ByVal ColIndex As Integer)
If StrOrdemNotas = " order by " & DataGrid2.Columns(ColIndex).DataField Then
  StrOrdemNotas = " order by " & DataGrid2.Columns(ColIndex).DataField & " desc"
Else
  StrOrdemNotas = " order by " & DataGrid2.Columns(ColIndex).DataField
End If

With dbClientesCobrancaTemp
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    Codigo = 0
  Else
    If IsNull(.Recordset!CodigoTemp) = False Then
      Codigo = .Recordset!CodigoTemp
    Else
      Codigo = 0
    End If
  End If
End With
With dbClientesNotaTemp
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from ClientesNotaTemp where codigoclientescobra=" & Codigo & StrOrdemNotas
  .Refresh
End With
End Sub

Private Sub dbClientesCobrancaTemp_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
With dbClientesCobrancaTemp
  If .Recordset.EOF = True Or .Recordset.BOF = True Then
    Codigo = 0
  Else
    If IsNull(.Recordset!CodigoTemp) = False Then
      Codigo = .Recordset!CodigoTemp
    Else
      Codigo = 0
    End If
  End If
End With
With dbClientesNotaTemp
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from ClientesNotaTemp where codigoclientescobra=" & Codigo & StrOrdemNotas
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
StrOrdemCobranca = " order by Nome"
StrOrdemNotas = " order by Cupom"
txtFechamento.Value = Date
With dbClientes
  .ConnectionString = CaminhoADO
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "Select *from clientes order by nome"
    Case ""
      .RecordSource = "Select *from clientes order by nome"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "Select *from clientes where planodeconta in (" & StrTemp & ") order by nome"
  End Select
  .Refresh
End With
With dbClientes2
  .ConnectionString = CaminhoADO
  Select Case Usuarios.Grupo.ClientesPlanos
    Case "0"
      .RecordSource = "Select *from clientes order by nome"
    Case ""
      .RecordSource = "Select *from clientes order by nome"
    Case Else
      StrTemp = "'" & Usuarios.Grupo.ClientesPlanos & "'"
      StrTemp = Replace(StrTemp, ",", "','")
      .RecordSource = "Select *from clientes where planodeconta in (" & StrTemp & ") order by nome"
  End Select
  .Refresh
End With

With dbClientesNotaTemp
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from ClientesNotaTemp where codigoclientescobra=0" & StrOrdemNotas
  .Refresh
End With
With dbClientesCobrancaTemp
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from ClientesCobrancaTemp where confirmado=0 and codigocliente=0" & StrOrdemCobranca
  .Refresh
End With
End Sub

Private Sub txtCodigo_Change()
txtCodigo2.Text = txtCodigo
End Sub

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbClientes
  .Refresh
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.Find "codigocliente=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    cboCliente.Text = .Recordset!Nome
    txtCodigo.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub txtCodigo2_GotFocus()
With txtCodigo2
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo2_LostFocus()
With dbClientes2
  .Refresh
  If IsNumeric(txtCodigo2.Text) = False Then Exit Sub
  .Recordset.Find "codigocliente=" & txtCodigo2.Text
  If .Recordset.EOF = False Then
    cboCliente2.Text = .Recordset!Nome
    txtCodigo2.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub txtFechamento_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtFechamento_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtFechamento_LostFocus()
Me.KeyPreview = True
End Sub
