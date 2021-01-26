VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliacaoCustodia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Custódia"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frmConciliaCustodia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbConcilia 
      Height          =   330
      Left            =   2040
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from concilianova"
      Caption         =   "dbConcilia"
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
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   2040
      Top             =   4080
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from contas"
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
   Begin MSAdodcLib.Adodc dbBloqueiaFechamento 
      Height          =   330
      Left            =   2040
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select *from bloqueiafechamento"
      Caption         =   "dbBloqueiaFechamento"
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
   Begin VB.Data qPendencias 
      Caption         =   "qPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from CompensaPendente where conciliado=0 and codigoconta=0"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data dbPendencias 
      Caption         =   "dbPendencias"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from CompensaPendente where conciliado=0 order by data"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirma 
      Caption         =   "Confirmar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4920
      Width           =   1095
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaCustodia.frx":0442
      Height          =   4335
      Left            =   120
      OleObjectBlob   =   "frmConciliaCustodia.frx":045D
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   56754177
      CurrentDate     =   37257
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   4560
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   3600
      TabIndex        =   5
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compensado em:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1230
   End
End
Attribute VB_Name = "frmConciliacaoCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoConta As Double

Private Sub cmdConfirma_Click()
Dim Resposta As Integer

With dbBloqueiaFechamento
  If .Recordset.EOF = False Then
    If .Recordset!Data1 <= txtData.Value And .Recordset!bloqueia1 = True Then
      MsgBox "Não pode ser feito este lançamento porque o fechamento está programado para " & .Recordset!Data1
      Exit Sub
    End If
  End If
End With

'If DateDiff("d", Date, txtData.Value) >= 1 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar custódia com data futura!"
'    Exit Sub
'  End If
'End If
'If DateDiff("d", Date, txtData.Value) <= -15 Then
'  If Usuarios.Grupo.AdmEstatus <> 2 Then
'    MsgBox "Somente usuário administrativo pode confirmar custódia com data anterior a 10 dias!"
'    Exit Sub
'  End If
'End If

With dbPendencias
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then
    MsgBox "Selecione uma custódia para confirmar!"
    Exit Sub
  End If
  Resposta = MsgBox("Os dados estão corretos?", vbYesNo + vbDefaultButton2)
  If Resposta = vbNo Then Exit Sub
  
  
  If dbContas.Recordset.EOF = True Or dbContas.Recordset.BOF = True Then
    MsgBox "Erro na tabela de contas!"
    Unload Me
    Exit Sub
  End If
  dbContas.Recordset.MoveFirst
  dbContas.Recordset.Find "codigoconta=" & CodigoConta
  If dbContas.Recordset.EOF = True Then
    MsgBox "Erro na tabela de contas!"
    Unload Me
    Exit Sub
  End If
  With dbConcilia
    .Recordset.AddNew
    .Recordset!CodigoConta = dbPendencias.Recordset!CodigoConta
    .Recordset!DataLanc = Now
    .Recordset!Data = txtData.Value
    .Recordset!compensado = True
    .Recordset!Tipo = "Custódia"
    .Recordset!Codigo = dbPendencias.Recordset!codigopendencia
    .Recordset!Descri = dbPendencias.Recordset!Descri
    .Recordset!NrDocumento = dbPendencias.Recordset!NrDoc
    .Recordset!Valor = dbPendencias.Recordset!Valor
    .Recordset.Update
  End With
  With dbContas
    .Recordset!Saldo = .Recordset!Saldo + dbPendencias.Recordset!Valor
    .Recordset.Update
  End With
  
  CompensaCustodia dbPendencias.Recordset!codigopendencia
  
  .Recordset.Edit
  .Recordset!conciliado = True
  .Recordset.Update
  .Refresh
End With

With qPendencias
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub dbPendencias_Reposition()
On Error Resume Next
If dbPendencias.Recordset.RecordCount = 0 Then Exit Sub
txtData.Value = dbPendencias.Recordset!Data
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case keysacii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
txtData.Value = Date
With dbConcilia
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select *from CompensaPendente where conciliado=0 and codigoconta=" & CodigoConta & " order by data"
  .Refresh
End With
With qPendencias
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as total from CompensaPendente where conciliado=0 and codigoconta=" & CodigoConta
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
With dbBloqueiaFechamento
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from bloqueiafechamento"
  .Refresh
End With

End Sub

Private Sub txtData_KeyDown(KeyCode As Integer, Shift As Integer)
Call Form_KeyPress(KeyCode)
End Sub
