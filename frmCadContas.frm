VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadContas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadatro de Contas"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmCadContas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   7365
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc qSaldos 
      Height          =   330
      Left            =   3000
      Top             =   1560
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
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
      RecordSource    =   "select sum(saldo) as total from contas"
      Caption         =   "qSaldos"
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
      Top             =   6345
      Width           =   7365
      _ExtentX        =   12991
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
      RecordSource    =   "select *from Contas order by descri"
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
      Bindings        =   "frmCadContas.frx":0442
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "CodigoConta"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0%"
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
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1725,165
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
      ScaleWidth      =   7365
      TabIndex        =   37
      Top             =   6015
      Width           =   7365
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   31
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   36
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   35
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   34
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Height          =   300
         Left            =   2408
         TabIndex        =   33
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   32
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   38
      Top             =   3120
      Width           =   7095
      Begin VB.ComboBox cboTipoConta 
         DataField       =   "Tipo"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2760
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Plano"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   7
         Left            =   120
         MaxLength       =   50
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "CodResumido"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0000000-0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   27
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Conta Interna"
         DataField       =   "Interna"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   5160
         TabIndex        =   30
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "CodigoFilial"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0000000-0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   3960
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "CodigoEmpresa"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0000000-0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   2760
         MaxLength       =   50
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Taxar CPMF"
         DataField       =   "TemCPMF"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "CC"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0000000-0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Agencia"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Banco"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Descri"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   120
         MaxLength       =   50
         TabIndex        =   4
         Top             =   480
         Width           =   3375
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         DataField       =   "Saldo"
         DataSource      =   "Adodc1"
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         DataField       =   "Previsao"
         DataSource      =   "Adodc1"
         Height          =   300
         Left            =   1920
         TabIndex        =   19
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         DataField       =   "Total"
         DataSource      =   "Adodc1"
         Height          =   300
         Left            =   3720
         TabIndex        =   21
         Top             =   1680
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox4 
         DataField       =   "Saldo"
         DataSource      =   "Adodc1"
         Height          =   300
         Left            =   5520
         TabIndex        =   23
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   "$#,##0.00;($#,##0.00)"
         PromptChar      =   " "
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Conta:"
         Height          =   195
         Index           =   12
         Left            =   2760
         TabIndex        =   28
         Top             =   2040
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Conta:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Resumido:"
         Height          =   195
         Index           =   10
         Left            =   1560
         TabIndex        =   26
         Top             =   2040
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Filial:"
         Height          =   195
         Index           =   9
         Left            =   3960
         TabIndex        =   13
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Empresa:"
         Height          =   195
         Index           =   8
         Left            =   2760
         TabIndex        =   11
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CPMF:"
         Height          =   195
         Index           =   7
         Left            =   5520
         TabIndex        =   22
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Index           =   6
         Left            =   3720
         TabIndex        =   20
         Top             =   1440
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Previsão:"
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   18
         Top             =   1440
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
         Height          =   195
         Index           =   3
         Left            =   1200
         TabIndex        =   9
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Banco:"
         Height          =   195
         Index           =   1
         Left            =   3600
         TabIndex        =   5
         Top             =   240
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      DataSource      =   "qSaldos"
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Saldo Total:"
      Height          =   195
      Left            =   4320
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmCadContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Adodc1_Reposition()
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition + 1
End Sub

Private Sub Adodc1_Validate(Action As Integer, Save As Integer)
If Save = True Then
  If QuerGravar = False Then
    Adodc1.Recordset.CancelUpdate
  End If
End If
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
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset


  Dim Resposta As Integer
  
  Resposta = MsgBox("Deseja excluir o registro atual?", vbYesNo, "Excluir!")
  If Resposta = vbNo Then
    Exit Sub
  End If
  
  With Adodc1.Recordset
    If .EOF = False Then
      If Adodc1.Recordset!Saldo <> 0 Then
        MsgBox "Esta conta possue saldo. Não será removida!"
        Exit Sub
      End If
      db.Open CaminhoADO
      dbTemp.Open "Select *from formadepagamento where codigoconta=" & Adodc1.Recordset!CodigoConta, db, adOpenStatic
      If dbTemp.RecordCount <> 0 Then
        MsgBox "Esta conta possue forma de pagamento direcionada para ela. Não será removida!"
        dbTemp.Close
        Exit Sub
      End If
      dbTemp.Close
      dbTemp.Open "Select *from cartoes where confirmado=0 and codigoconta=" & Adodc1.Recordset!CodigoConta, db, adOpenStatic
      If dbTemp.RecordCount <> 0 Then
        MsgBox "Esta conta possue cartão pendente direcionado para ela. Não será removida!"
        dbTemp.Close
        Exit Sub
      End If
      dbTemp.Close
      db.Close
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

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  qSaldos.Refresh
  Frame1.Enabled = False
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
With qSaldos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With cboTipoConta
  .Clear
  .AddItem "Outros"
  .AddItem "Produto"
  .AddItem "Cliente"
  .AddItem "Despesa"
  .AddItem "Cheques"
  .AddItem "Cartão"
  .AddItem "Fornecedores"
End With

Select Case Usuarios.Grupo.CadConta
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

