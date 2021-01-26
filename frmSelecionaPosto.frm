VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSelecionaPosto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleciona Posto"
   ClientHeight    =   3975
   ClientLeft      =   3435
   ClientTop       =   1860
   ClientWidth     =   5070
   Icon            =   "frmSelecionaPosto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc dbPostoInfo 
      Height          =   330
      Left            =   1440
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=postoinfo.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=postoinfo.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from usuarios"
      Caption         =   "dbPostoInfo"
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
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfigura 
      Caption         =   "Configurar"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "Selecionar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Banco de Dados"
      FileName        =   "*.mdb"
      Filter          =   "Banco de dados |*.mdb"
   End
   Begin MSDataListLib.DataList lstPostos 
      Bindings        =   "frmSelecionaPosto.frx":0442
      Height          =   3180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5609
      _Version        =   393216
      ListField       =   "Nome"
      BoundColumn     =   ""
      Object.DataMember      =   ""
   End
End
Attribute VB_Name = "frmSelecionaPosto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfigura_Click()
  Dim Arquivo As String
  With dbPostoInfo
    If .Recordset.RecordCount = 0 Then Exit Sub
    If .Recordset.EOF = True Then Exit Sub
    If lstPostos.Text <> .Recordset!Nome Then
      MsgBox "Erro na listagem de postos!"
      Exit Sub
    End If
    With CommonDialog1
      On Error GoTo Cancelado
      .ShowOpen
      Arquivo = .filename
    End With
    On Error GoTo 0
    StrTemp = InputBox("Informe um novo nome!", "Configurar", lstPostos.Text)
    If StrTemp = "" Then GoTo Cancelado
    .Recordset!Nome = StrTemp
    .Recordset!Dados = Arquivo
    .Recordset.Update
    .Recordset.Requery
  End With
  Exit Sub
Cancelado:
  MsgBox "Não foi possível definir o banco de dados! Posto não configurado!"
  
End Sub

Private Sub cmdSair_Click()
End
End Sub

Private Sub cmdSelecionar_Click()
Dim StrTemp As String
With dbPostoInfo
  If .Recordset.EOF = True Then Exit Sub
  Caminho = .Recordset!Dados
  NomePosto = .Recordset!Nome
  'esse controle passou para a tabela Postos
  ComissaoAcumulativa = .Recordset!ComissaoAcumulativa
  
  On Error Resume Next
  StrTemp = Dir(Caminho)
  If Err.Number <> 0 Then
    MsgBox "O caminho especificado pode não estar acessível."
    Exit Sub
  End If
  On Error GoTo 0
  For i = Len(Caminho) To 1 Step -1
    If Mid(Caminho, i, 1) = "\" Then
      Diretorio = Mid(Caminho, 1, i)
      Exit For
    End If
  Next i
End With
Unload Me
If Selecionando = True Then
  If Provedor = "SQLOLEDB.1" Then
    CaminhoADO = "Provider=" & Provedor & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Maria Vitoria;Data Source=temvale17"
  Else
    CaminhoADO = "Provider=" & Provedor & ";Data Source=" & Caminho & ";Persist Security Info=False"
  End If
End If

Unload Me
End Sub

Private Sub dbPostos_Reposition()
On Error Resume Next
txtNome.Text = dbPostos.Recordset!Nome
End Sub

Private Sub Form_Load()

With dbPostoInfo
  .ConnectionString = "Provider=" & Provedor & ";Data Source=" & App.Path & "\PostoInfo.mdb;Persist Security Info=False"
  .Refresh
  CaminhoUsuarios = .Recordset!Caminho
  If Provedor = "SQLOLEDB.1" Then
    CaminhoUsuariosAdo = "Provider=" & Provedor & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Usuarios;Data Source=temvale17"
  Else
    CaminhoUsuariosAdo = "Provider=" & Provedor & ";Data Source=" & .Recordset!Caminho & ";Persist Security Info=False"
  End If
  On Error GoTo TrataErro
  .ConnectionString = CaminhoUsuariosAdo
  .RecordSource = "select *from caminhonovo order by nome"
  .Refresh
End With

Exit Sub

TrataErro:
With CommonDialog1
  .ShowOpen
  CaminhoUsuarios = .filename
End With

With dbPostoInfo
  .ConnectionString = "Provider=" & Provedor & ";Data Source=" & App.Path & "\PostoInfo.mdb;Persist Security Info=False"
  .RecordSource = "select *from usuarios"
  .Refresh
  .Recordset!Caminho = CaminhoUsuarios
  .Recordset.Update
  
  If Provedor = "SQLOLEDB.1" Then
    CaminhoUsuariosAdo = "Provider=" & Provedor & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Usuarios;Data Source=temvale17"
  Else
    CaminhoUsuariosAdo = "Provider=" & Provedor & ";Data Source=" & .Recordset!Caminho & ";Persist Security Info=False"
  End If
  On Error GoTo TrataErro
  .ConnectionString = CaminhoUsuariosAdo
  .RecordSource = "select *from caminhonovo order by nome"
  .Refresh
End With


End Sub


Private Sub lstPostos_Click()
With dbPostoInfo
  On Error Resume Next
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & lstPostos.Text & "'"
End With
End Sub

Private Sub lstPostos_DblClick()
Call lstPostos_Click
Call cmdSelecionar_Click
End Sub
