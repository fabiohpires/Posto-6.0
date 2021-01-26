VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadVendedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Funcionários"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   Icon            =   "frmCadVendedor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   3360
      Top             =   2040
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
   Begin MSAdodcLib.Adodc dbDespesaTipo 
      Height          =   330
      Left            =   3240
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "select *from despesatipo where descri like 'Func*' order by descri"
      Caption         =   "dbDespesaTipo"
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
      Top             =   5490
      Width           =   7005
      _ExtentX        =   12356
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
      RecordSource    =   "select *from Vendedores order by nome"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3240
      Top             =   1680
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
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
      RecordSource    =   "select *from turnos order by descri"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6000
      Picture         =   "frmCadVendedor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "Imprimir"
      Top             =   2400
      Width           =   735
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      ScaleHeight     =   330
      ScaleWidth      =   7005
      TabIndex        =   19
      Top             =   5160
      Width           =   7005
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Adicionar"
         Height          =   300
         Left            =   1313
         TabIndex        =   14
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remover"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2408
         TabIndex        =   15
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Atuali&zar"
         Height          =   300
         Left            =   3503
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Gravar"
         Height          =   300
         Left            =   4598
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Cancel          =   -1  'True
         Caption         =   "&Fechar"
         Height          =   300
         Left            =   5693
         TabIndex        =   18
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   300
         Left            =   233
         TabIndex        =   13
         Top             =   0
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmCadVendedor.frx":0EC4
      Height          =   2175
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
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
         DataField       =   "Gerente"
         Caption         =   "Caixa"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   5
            Format          =   ""
            HaveTrueFalseNull=   1
            TrueValue       =   "Sim"
            FalseValue      =   "Não"
            NullValue       =   ""
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   7
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3869,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   854,929
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   5775
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmCadVendedor.frx":0ED9
         DataField       =   "CodigoTurno"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2520
         TabIndex        =   7
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoTurno"
         Text            =   ""
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Caixa"
         DataField       =   "Gerente"
         DataSource      =   "Adodc1"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Funcao"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Codigo"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Nome"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmCadVendedor.frx":0EEE
         DataField       =   "codigoDespesa"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoDespesa"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo5 
         Bindings        =   "frmCadVendedor.frx":0F0A
         DataField       =   "PlanoDeConta"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Descri"
         BoundColumn     =   "CodigoConta"
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Plano de Conta:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Despesa Vinculada:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Função:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Turno:"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmCadVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String
Dim CodigoFuncionario As String
Dim Adicionando As Boolean

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 16

StrTemp = "Relação de Funcionários"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 14



StrTemp = "Data: " & Format(Dia, "Long Date")
Printer.CurrentX = 0
Printer.Print StrTemp;
StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp



StrTemp = "Código"
Printer.CurrentX = 20 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Nome"
Printer.CurrentX = 23
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 140
Printer.Print StrTemp;

StrTemp = "Caixa"
Printer.CurrentX = 170
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1




End Sub



Private Sub Adodc1_Reposition()
Adodc1.Caption = "Registro: " & Adodc1.Recordset.AbsolutePosition
End Sub

Private Sub Adodc1_Validate(Action As Integer, Save As Integer)
If Save = True Then
  If QuerGravar = False Then
    Adodc1.Recordset.CancelUpdate
  End If
End If
End Sub

Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
CodigoFuncionario = 0
If Adicionando = True Then Exit Sub
If IsNumeric(txtFields(1).Text) = True Then
  CodigoFuncionario = txtFields(1).Text
End If
End Sub

Private Sub cmdAdd_Click()
  Adicionando = True
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
      
      'verifica se existe alguma pendencia para pagar ou receber
      
      db.Open CaminhoADO
      dbTemp.Open "Select *from vales where cobrado=0 and codfun=" & Adodc1.Recordset!codigovendedor, db
      If dbTemp.RecordCount > 0 Then
        MsgBox "Existe vale não cobrado deste funcionário em aberto!"
        Exit Sub
      End If
      dbTemp.Close
      dbTemp.Open "Select *from venda2 where pago=0 and codigovendedor=" & Adodc1.Recordset!codigovendedor, db
      If dbTemp.RecordCount > 0 Then
        MsgBox "Existe comissão não paga deste funcionário!"
        Exit Sub
      End If
      
      
      .Delete
      If .EOF = False Then
      .MoveNext
      Else
        If .BOF = False Then .MoveLast
      End If
    End If
  End With
  
  Frame1.Enabled = False
  dbTemp.Close
  db.Close
End Sub

Private Sub cmdEditar_Click()
If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Frame1.Enabled = True
txtFields(1).SetFocus
'If Usuarios.Nome = "Usuário Master" Then
'  txtFields(1).Enabled = True
'  txtFields(1).SetFocus
'Else
'  txtFields(1).Enabled = False
'  txtFields(0).SetFocus
'End If
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double

With Adodc1
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Largura = 190
  Dia = Now
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      Printer.CurrentY = 0
      Printer.NewPage
    End If
    
    If IsNull(.Recordset!Codigo) = False Then
      StrTemp = .Recordset!Codigo
      Printer.CurrentX = 20 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
    End If
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 23
    Printer.Print StrTemp;
    
    Adodc2.Refresh
    If Adodc2.Recordset.RecordCount <> 0 Then
      Adodc2.Recordset.MoveFirst
      Adodc2.Recordset.Find "codigoturno=" & .Recordset!CodigoTurno
      If Adodc2.Recordset.EOF = False Then
        StrTemp = Adodc2.Recordset!Descri
        Printer.CurrentX = 140
        Printer.Print StrTemp;
      End If
    End If
    If .Recordset!gerente = True Then
      StrTemp = "Sim"
    Else
      StrTemp = "Não"
    End If
    Printer.CurrentX = 170
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  .Recordset.MoveFirst
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.EndDoc
End With

NaoImprime:

End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  Adodc1.Refresh
  Frame1.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
  'On Error Resume Next
  With Adodc1
    If CodigoFuncionario <> 0 And .Recordset!Codigo <> CodigoFuncionario Then
      On Error Resume Next
      If AtualizaCodigoFuncionario(.Recordset!Codigo, .Recordset!codigovendedor, CodigoFuncionario) = False Then
        txtFields(1).Text = CodigoFuncionario
      End If
    End If
    A = .Recordset.AbsolutePosition
    .Recordset.Update
    .Recordset.AbsolutePosition = A
    
  End With
  
  cmdAdd.Enabled = True
  cmdDelete.Enabled = True
  cmdRefresh.Enabled = True
  Frame1.Enabled = False
  Adicionando = False
End Sub

Private Sub cmdClose_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Adodc1.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
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
strOrdem = " order by Nome"
Adicionando = False
With Adodc1
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from Vendedores" & strOrdem
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
  End If
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbDespesaTipo
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from despesatipo order by descri"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "As despesas relativos a funcionários devem iniciar com 'Func'!"
  End If
End With
With Adodc2
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from turnos order by descri"
  .Refresh
End With
Select Case Usuarios.Grupo.CadFuncionarios
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


