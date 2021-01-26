VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmAlteraPreco 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alteração de preços de Combustível"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   Icon            =   "frmAdmAlteraPreco.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin MSDataListLib.DataCombo cboTurno 
      Bindings        =   "frmAdmAlteraPreco.frx":0442
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbFechamento 
      Height          =   330
      Left            =   5160
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "fechamentodecaixa"
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
   Begin MSAdodcLib.Adodc dbBicos 
      Height          =   330
      Left            =   5160
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "bicos"
      Caption         =   "dbBicos"
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
   Begin MSAdodcLib.Adodc dbAlteraBico 
      Height          =   330
      Left            =   5160
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "alterabico"
      Caption         =   "dbAlteraBico"
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
   Begin MSAdodcLib.Adodc dbTurnos 
      Height          =   330
      Left            =   5160
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from Turnos order by descri"
      Caption         =   "dbTurnos"
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
   Begin MSAdodcLib.Adodc dbAlteracao 
      Height          =   330
      Left            =   5160
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=Posto.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from Alteracoes order by dataalteracao, turno"
      Caption         =   "dbAlteracao"
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
      Bindings        =   "frmAdmAlteraPreco.frx":0459
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6800
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "DataAlteracao"
         Caption         =   "Data Alteração"
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
      BeginProperty Column01 
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   1409,953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1574,929
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRemover 
      Caption         =   "Remover"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   529
      _Version        =   393216
      Format          =   132775937
      CurrentDate     =   38715
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmAdmAlteraPreco.frx":0473
      Height          =   3855
      Left            =   3840
      TabIndex        =   8
      Top             =   720
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Bico"
         Caption         =   "Bico"
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
         DataField       =   "Produto"
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
         DataField       =   "Preco"
         Caption         =   "Preço"
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
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   510,236
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2250,142
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1080
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Dia a alterar:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdmAlteraPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboTurno_LostFocus()
With dbTurnos
  .Refresh
  If cboTurno.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboTurno.Text & "'"
  If .Recordset.EOF = False Then
    cboTurno.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdNovo_Click()
Dim CodigoAlteracao As Double, DataCaixa As Date, CodigoFechamento As Double

If Usuarios.Nome <> "Usuário Master" Then
  If DateDiff("d", Date, txtData.Value) < -30 Then
    MsgBox "Data muito antiga!"
    Exit Sub
  End If
  If DateDiff("d", Date, txtData.Value) > 30 Then
    MsgBox "Data muito futura!"
    Exit Sub
  End If
End If
If cboTurno.Text = "" Then
  MsgBox "Indique um turno!"
  cboTurno.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset.EOF = True Then
  MsgBox "Indique um turno!"
  cboTurno.SetFocus
  Exit Sub
End If
If dbTurnos.Recordset!Descri <> cboTurno.Text Then
  MsgBox "Turno incorreto!"
  cboTurno.SetFocus
  Exit Sub
End If
With dbFechamento
  .RecordSource = "Select *from fechamentodecaixa where datacaixa=#" & DataInglesa(txtData.Value) & "# and codigoturno=" & dbTurnos.Recordset!CodigoTurno
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If .Recordset!fechado = True Then
      MsgBox "Este caixa já foi finalizado! Não será possível fazer a alteração de preço!"
      Exit Sub
    End If
  End If
End With
With dbAlteracao
  StrTemp = .RecordSource
  .RecordSource = "select *from alteracoes where dataalteracao=#" & DataInglesa(txtData.Value) & "# and codigoturno=" & dbTurnos.Recordset!CodigoTurno
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.EOF = False Then
      MsgBox "Alteração já criada!"
      Exit Sub
    Else
      .Recordset.AddNew
    End If
  Else
    .Recordset.AddNew
  End If
  .Recordset!dataalteracao = txtData.Value
  .Recordset!datacriada = Now
  .Recordset!CodigoTurno = dbTurnos.Recordset!CodigoTurno
  .Recordset!Turno = dbTurnos.Recordset!Descri
  .Recordset.Update
  .Refresh
  CodigoAlteracao = .Recordset!codalteracao
  .RecordSource = StrTemp
  .Refresh
  .Recordset.Find "codalteracao=" & CodigoAlteracao
End With
With dbAlteraBico
  .Refresh
  If .Recordset.RecordCount = 0 Then
    With dbBicos
      .RecordSource = "select bicos.*, produtos.* from bicos, produtos where bicos.codigoproduto=produtos.codigoproduto"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        .Recordset.MoveLast
        .Recordset.MoveFirst
        DataCaixa = CDate(txtData.Value & " " & dbTurnos.Recordset!HoraIni)
        Do While .Recordset.EOF = False
          dbAlteraBico.Recordset.AddNew
          dbAlteraBico.Recordset!codalteracao = CodigoAlteracao
          dbAlteraBico.Recordset!Bico = .Recordset!Bico
          dbAlteraBico.Recordset!Produto = .Recordset!Descri
          dbAlteraBico.Recordset!Preco = 0
          dbAlteraBico.Recordset.Update
          .Recordset.MoveNext
        Loop
      End If
      StrTemp = .RecordSource
      .RecordSource = "select *from alteracoes order by dataalteracao, turno"
      .Refresh
      .Recordset.Find "codalteracao=" & CodigoAlteracao
      If .Recordset.EOF = True Then
        GoTo Termina
      End If
      .Recordset.MovePrevious
      A = .Recordset!codalteracao
      .RecordSource = "select *from alterabico where codalteracao=" & A & " order by bico"
      .Refresh
      If .Recordset.RecordCount <> 0 Then
        Do While .Recordset.EOF = False
          dbAlteraBico.Refresh
          dbAlteraBico.Recordset.MoveFirst
          dbAlteraBico.Recordset.Find "bico=" & .Recordset!Bico
          If dbAlteraBico.Recordset.EOF = False Then
            dbAlteraBico.Recordset!Preco = .Recordset!Preco
            dbAlteraBico.Recordset.Update
          End If
          .Recordset.MoveNext
        Loop
      End If
      .RecordSource = StrTemp
      .Refresh
    End With

  End If
End With
Termina:
dbAlteraBico.Refresh

End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
If dbAlteracao.Recordset.EOF = True Then Exit Sub
Resposta = MsgBox("Deseja remover a alteracao atual?", vbYesNo + vbDefaultButton2)
If Resposta = vbNo Then Exit Sub
With dbAlteraBico
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    Do While .Recordset.EOF = False
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
dbAlteracao.Recordset.Delete
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbAlteracao_Reposition()

End Sub

Private Sub DataGrid2_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub DataGrid2_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub DBGrid1_Click()

End Sub

Private Sub dbAlteracao_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If dbAlteracao.Recordset.EOF = True Then Exit Sub
If dbAlteracao.Recordset.BOF = True Then Exit Sub
With dbAlteraBico
  .ConnectionString = CaminhoADO
  If IsNull(dbAlteracao.Recordset!codalteracao) = False Then
    CodigoAlteracao = dbAlteracao.Recordset!codalteracao
  Else
    CodigoAlteracao = 0
  End If
  .RecordSource = "select *from alterabico where codalteracao=" & CodigoAlteracao & " order by bico"
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
txtData.Value = Date
With dbAlteracao
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from alteracoes"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
  End If
End With
With dbTurnos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbBicos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbAlteraBico
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbFechamento
  .ConnectionString = CaminhoADO
  .Refresh
End With
If dbAlteracao.Recordset.RecordCount <> 0 Then dbAlteracao.Recordset.MoveLast

Select Case Usuarios.Grupo.AdmEstatus
  Case 1 'Somente leitura
    cmdNovo.Enabled = False
    cmdRemover.Enabled = False
  Case 2 'Liberado
    
End Select
End Sub

Private Sub txtData_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtData_LostFocus()
Me.KeyPreview = True
End Sub
