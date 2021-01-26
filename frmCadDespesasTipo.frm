VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCadDespesasTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Tipo de Despesa"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   Icon            =   "frmCadDespesasTipo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   1560
      Top             =   1320
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      RecordSource    =   "select *from contas order by descri"
      Caption         =   "dbContas"
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
   Begin MSAdodcLib.Adodc dbDespesaTipoObrigatorio 
      Height          =   330
      Left            =   1560
      Top             =   960
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
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
      RecordSource    =   "select *from despesatipoobrigatorio"
      Caption         =   "dbDespesaTipoObrigatorio"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6495
      Width           =   6990
      _ExtentX        =   12330
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
      RecordSource    =   "select *from DespesaTipo order by descri"
      Caption         =   "Adodc1"
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
      Bindings        =   "frmCadDespesasTipo.frx":0442
      Height          =   2055
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3625
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
      BeginProperty Column01 
         DataField       =   "CodigoNoPosto"
         Caption         =   "Código no Posto"
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
            ColumnWidth     =   4305,26
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1695,118
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   6990
      TabIndex        =   24
      Top             =   6165
      Width           =   6990
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   19
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2408
         TabIndex        =   20
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   21
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   22
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   23
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   25
      Top             =   2160
      Width           =   6735
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmCadDespesasTipo.frx":0457
         DataField       =   "CodigoPlanoDeConta"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   1680
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoConta"
         Text            =   ""
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CodigoHistorico"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1680
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc dbDespesaTipoGrupo 
         Height          =   330
         Left            =   1920
         Top             =   3120
         Visible         =   0   'False
         Width           =   3495
         _ExtentX        =   6165
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
         RecordSource    =   "DespesaTipoSubGrupo"
         Caption         =   "dbDespesaTipoGrupo"
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
      Begin VB.CheckBox Check1 
         Caption         =   "Despesa com Vencimento Mensal"
         DataField       =   "Mensal"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   2040
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         Format          =   72351745
         CurrentDate     =   39468
      End
      Begin VB.CommandButton cmdRemover 
         Caption         =   "Remover"
         Height          =   375
         Left            =   5520
         TabIndex        =   16
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   4440
         TabIndex        =   15
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtDescri 
         Height          =   285
         Left            =   120
         MaxLength       =   50
         TabIndex        =   14
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtFields 
         DataField       =   "CodigoNoPosto"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Descri"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   1
         Top             =   480
         Width           =   3735
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Bindings        =   "frmCadDespesasTipo.frx":046E
         Height          =   1095
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   1931
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
         ColumnCount     =   1
         BeginProperty Column00 
            DataField       =   "Descri"
            Caption         =   "Sub-Categoria"
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
            BeginProperty Column00 
               ColumnWidth     =   5820,095
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo ctlDescri 
         Bindings        =   "frmCadDespesasTipo.frx":048F
         DataField       =   "Obrigatorio"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   3960
         TabIndex        =   3
         Top             =   480
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "Descri"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Conta:"
         Height          =   195
         Index           =   3
         Left            =   1560
         TabIndex        =   11
         Top             =   1440
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Histórico:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label4 
         Caption         =   "Obrigatório:"
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Último vencimento:"
         Height          =   255
         Left            =   2040
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Sub-Categoria:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código no Posto:"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmCadDespesasTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_Validate(Action As Integer, Save As Integer)
If Save = True Then
  If QuerGravar = False Then
    Adodc1.Recordset.CancelUpdate
  End If
End If

End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Dim CodigoDespesa As Double

Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1

If Adodc1.Recordset.EOF = False Then
  If IsNull(Adodc1.Recordset!CodigoDespesa) = False Then
    CodigoDespesa = Adodc1.Recordset!CodigoDespesa
  Else
    CodigoDespesa = 0
  End If
Else
  CodigoDespesa = 0
End If
With dbDespesaTipoGrupo
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from despesatiposubgrupo where codigodespesatipo=" & CodigoDespesa & " order by descri"
  .Refresh
End With
End Sub

Private Sub cmdAdd_Click()
  Adodc1.Recordset.AddNew
  cmdAdd.Enabled = False
  cmdDelete.Enabled = False
  cmdRefresh.Enabled = False
  Frame1.Enabled = True
  txtFields(0).SetFocus
End Sub

Private Sub cmdDelete_Click()
  Dim Resposta As Integer
  
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo, "Excluir!")
  If Resposta = vbNo Then
    Exit Sub
  End If
  
  With Adodc1.Recordset
    If .EOF = False Then
      .Delete
      If .EOF = False Then
      .MoveNext
      Else
        If .BOF = False Then .MoveLast
      End If
    End If
  End With
  
  Frame1.Enabled = False
End Sub

Private Sub cmdEditar_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Frame1.Enabled = True
txtFields(0).SetFocus
End Sub

Private Sub cmdIncluir_Click()
Dim CodigoDespesa As Double
If Adodc1.Recordset.EOF = True Then
  MsgBox "Selecione uma despesa primeiro!"
  Exit Sub
End If
CodigoDespesa = Adodc1.Recordset!CodigoDespesa

With dbDespesaTipoGrupo
  .Recordset.AddNew
  .Recordset!codigodespesatipo = CodigoDespesa
  .Recordset!Descri = txtDescri.Text
  .Recordset.Update
End With
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdRemover_Click()
Dim Resposta As Integer
With dbDespesaTipoGrupo
  If .Recordset.EOF = True Then
    MsgBox "Escolha uma Sub-Categoria para remover!"
    Exit Sub
  End If
  Resposta = MsgBox("Deseja remover a Sub-Categoria atual?", vbYesNo)
  If Resposta = vbNo Then Exit Sub
  .Recordset.Delete
  .Refresh
End With

End Sub

Private Sub cmdUpdate_Click()
  On Error Resume Next
  With Adodc1
    A = .Recordset.AbsolutePosition
    .Recordset.Update
    .Recordset.AbsolutePosition = A
  End With
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Frame1.Enabled = False

End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select

End Sub

Private Sub Form_Load()
With Adodc1
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbDespesaTipoGrupo
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbDespesaTipoObrigatorio
  .ConnectionString = CaminhoADO
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .Recordset.AddNew
    .Recordset!Descri = "Observações"
    .Recordset.Update
    .Recordset.AddNew
    .Recordset!Descri = "Mes e Ano Referência"
    .Recordset.Update
    .Recordset.AddNew
    .Recordset!Descri = "Período"
    .Recordset.Update
    .Recordset.AddNew
    .Recordset!Descri = "Obs. Adicional"
    .Recordset.Update
    .Recordset.AddNew
    .Recordset!Descri = "Nenhum"
    .Recordset.Update
  End If
End With

Select Case Usuarios.Grupo.CadDespesaBancaria
  Case 1 'Somente leitura
    cmdEditar.Enabled = False
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

