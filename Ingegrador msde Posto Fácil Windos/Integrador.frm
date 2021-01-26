VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmIntegrador 
   Caption         =   "Integrador p/ Posto Fácil"
   ClientHeight    =   2970
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4485
   Icon            =   "Integrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
      Begin MSAdodcLib.Adodc DbEncerrantes 
         Height          =   375
         Left            =   120
         Top             =   1440
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "DbEncerrantes"
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
      Begin MSAdodcLib.Adodc dbDestino 
         Height          =   330
         Left            =   240
         Top             =   720
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
         UserName        =   "sa"
         Password        =   "masterkey"
         RecordSource    =   ""
         Caption         =   "dbDestino"
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
      Begin MSAdodcLib.Adodc dbCaixa 
         Height          =   330
         Left            =   240
         Top             =   360
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
         UserName        =   "sa"
         Password        =   "masterkey"
         RecordSource    =   ""
         Caption         =   "dbCaixa"
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
      Begin MSAdodcLib.Adodc storExcluirCaixa 
         Height          =   330
         Left            =   240
         Top             =   1080
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
         CommandType     =   4
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
         Connect         =   "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=temvale17"
         OLEDBString     =   "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=temvale17"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "sa"
         Password        =   "masterkey"
         RecordSource    =   "spApagaCaixa;1"
         Caption         =   "storExcluirCaixa"
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
      Begin MSAdodcLib.Adodc dbVendas 
         Height          =   375
         Left            =   120
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbVendas"
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
         Left            =   120
         Top             =   2160
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
      Begin MSAdodcLib.Adodc dbNumerarios 
         Height          =   375
         Left            =   120
         Top             =   2880
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbNumerarios"
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
      Begin MSAdodcLib.Adodc dbInventario 
         Height          =   375
         Left            =   120
         Top             =   2520
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbInventario"
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
      Begin MSAdodcLib.Adodc dbPessoa 
         Height          =   375
         Left            =   120
         Top             =   3240
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbPessoa"
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
      Begin MSAdodcLib.Adodc dbDespesas 
         Height          =   375
         Left            =   120
         Top             =   3600
         Width           =   2775
         _ExtentX        =   4895
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
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "dbDespesas"
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
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Configurar 
      Caption         =   "Configurar"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      _Version        =   393216
      FullWidth       =   281
      FullHeight      =   81
   End
   Begin VB.ComboBox cboTurno 
      Height          =   315
      ItemData        =   "Integrador.frx":0442
      Left            =   1800
      List            =   "Integrador.frx":0452
      TabIndex        =   3
      Text            =   "01"
      Top             =   360
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   2400
      Top             =   840
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   1680
      Top             =   840
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtDataCaixa 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   114163713
      CurrentDate     =   39034
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Data Inicial:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Turno:"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmIntegrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoPosto As String, NomePosto As String
Public Caminho As String, CaminhoFirebird, Conectar As String, Destino As String
Public WsPrincipal As Workspace, DbPrincipal As Database
Public strMSDE As String, IniciaComWindows As Boolean
Public strDBase As String

Public FbUsuario As String
Public FbSenha As String

Public Function GravarConfiguracoes(ByVal strOrigem As String, ByVal strDestino As String, ByVal dtDataCaixa As String, ByVal strTurno As String, ByVal BancoDeDados As String, ByVal strCodigoPosto As String, ByVal strNomePosto As String) As Boolean
GravarConfiguracoes = False
SaveSetting App.EXEName, "Config", "Codigo", strCodigoPosto
SaveSetting App.EXEName, "Config", "Nome", strNomePosto
SaveSetting App.EXEName, "Config", "Origem", strOrigem
SaveSetting App.EXEName, "Config", "Destino", strDestino
SaveSetting App.EXEName, "Config", "Ultimo Dia", dtDataCaixa
SaveSetting App.EXEName, "Config", "Ultimo Turno", strTurno
SaveSetting App.EXEName, "Config", "MSDE", BancoDeDados
'firebird
'Provider=MSDASQL.1;Persist Security Info=False;Extended Properties="DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; DBNAME=c:\rede\PFWIN - Cópia.GDB;"
CaminhoFirebird = "Provider=MSDASQL.1;Persist Security Info=False;Extended Properties=""DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; DBNAME=" & strOrigem & ";"""

SaveSetting App.EXEName, "Config", "CaminhoFireBird", CaminhoFirebird
Caminho = strOrigem
Destino = strDestino
strMSDE = BancoDeDados
CodigoPosto = strCodigoPosto
NomePosto = strNomePosto


GravarConfiguracoes = True
End Function

Public Function PegaConfiguracoes() As Boolean
CodigoPosto = GetSetting(App.EXEName, "Config", "Codigo")
NomePosto = GetSetting(App.EXEName, "Config", "Nome")
Caminho = GetSetting(App.EXEName, "Config", "Origem")
Destino = GetSetting(App.EXEName, "Config", "Destino")
strMSDE = GetSetting(App.EXEName, "Config", "MSDE", "SQLOLEDB.1")
cboTurno.Text = GetSetting(App.EXEName, "Config", "Ultimo Turno", "01")
txtDataCaixa.Value = CDate(GetSetting(App.EXEName, "Config", "Ultimo Dia", Date))
CaminhoFirebird = GetSetting(App.EXEName, "Config", "CaminhoFireBird")

Ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador")
If Ret = "" Then
  IniciaComWindows = False
Else
  IniciaComWindows = True
End If

End Function

Private Sub PegaEncerrantes(ByVal Dia As Date, ByVal Turno As String)
Dim strEncerrantes As String, IntArquivo As Integer
Dim StrTemp As String, StrLinha As String
Dim StrDb As String

TentaDeNovo:
On Error GoTo 0
'On Error GoTo trataErro
lblStatus.Caption = "Verificando se existe registros já exportados..."
lblStatus.Refresh
With storExcluirCaixa
  .ConnectionString = "Provider=" & strMSDE & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & Destino
  .RecordSource = "spApagaCaixa;1('" & Dia & "','" & Turno & "','" & CodigoPosto & "')"
  On Error Resume Next
  .Refresh
  If Err.Number <> 0 Then
    Resposta = MsgBox("Não foi possível conectar com o servidor! Deseja tentar novamente?", vbYesNo)
    If Resposta = vbYes Then
      GoTo TentaDeNovo
    Else
      lblStatus.Caption = ""
      Exit Sub
    End If
  End If
End With

lblStatus.Caption = "Carregando Tabela de Caixas..."
lblStatus.Refresh

With dbDestino
  StrDb = "Integrador"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & StrDb & " ;Data Source=" & Destino
  .ConnectionString = Conectar
  .RecordSource = "Select *from caixas where codigoposto='" & CodigoPosto & "' and datacaixa='" & Dia & "' and turno='" & Turno & "'"
  .Refresh
End With
On Error GoTo 0
lblStatus.Caption = "Carregando Tabela de Bicos..."
lblStatus.Refresh

'On Error GoTo TrataErro
With DbEncerrantes
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select * from movimento_enc where dat_movimento='" & DataInglesa(Dia) & "' and turno_id='" & CInt(Turno) & "'"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Vendas..."
lblStatus.Refresh
With dbVendas
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select movimento_produto.*, pessoa.*  from movimento_produto, pessoa where movimento_produto.pessoa_id=pessoa.pessoa_id and dat_movimento='" & DataInglesa(Dia) & "' and turno_id=" & CInt(Turno)
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Clientes..."
lblStatus.Refresh
With dbClientes
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select ecf_consumidor.*, item_ecf_consumidor.* from ecf_consumidor, item_ecf_consumidor where item_ecf_consumidor.ecf_consumidor_id=ecf_consumidor.ecf_consumidor_id and dat_movimento='" & DataInglesa(Dia) & "' and turno_id=" & CInt(Turno) & " and pessoa_id is not null"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Inventários..."
lblStatus.Refresh
With dbInventario
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "select *from medicao_fisica where dat_movimento='" & DataInglesa(Dia) & "' and turno_id=" & CInt(Turno)
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Numerários..."
lblStatus.Refresh
With dbNumerarios
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select dat_movimento, turno_id, cartao_id, sum(val_Cartao) as total from movimento_cartao where dat_movimento='" & DataInglesa(Dia) & "' and turno_id=" & CInt(Turno) & " group by cartao_id, dat_movimento, turno_id"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Funcionários..."
lblStatus.Refresh
With dbPessoa
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select *from pessoa"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Despesas..."
lblStatus.Refresh
With dbDespesas
  .ConnectionString = CaminhoFirebird
  .UserName = FbUsuario
  .Password = FbSenha
  .RecordSource = "Select *from despesa_Caixa where dat_movimento='" & DataInglesa(Dia) & "' and turno_id=" & CInt(Turno)
  .Refresh
End With

With dbDestino
  .Recordset.AddNew
  .Recordset!DataCaixa = Dia
  .Recordset!Turno = Turno
  .Recordset!CodigoPosto = CodigoPosto
  .Recordset!NomePosto = NomePosto
  .Recordset!linhaexportada = "000|Inicio do Caixa"
  .Recordset.Update
End With


If DbEncerrantes.Recordset.RecordCount <> 0 Then
  DbEncerrantes.Recordset.MoveLast
  DbEncerrantes.Recordset.MoveFirst
  Do While DbEncerrantes.Recordset.EOF = False
    lblStatus.Caption = "Exportando Bicos... " & Format((DbEncerrantes.Recordset.AbsolutePosition / (DbEncerrantes.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    StrTemp = Space(6)
    If IsNull(DbEncerrantes.Recordset!bico_id) = False Then
      If DbEncerrantes.Recordset!bico_id <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!bico_id) + 1) = DbEncerrantes.Recordset!bico_id
      End If
    End If
    StrLinha = "001|" & StrTemp
    
    StrTemp = Space(16)
'    If IsNull(DbEncerrantes.Recordset!d07_encabe) = False Then
'      If DbEncerrantes.Recordset!d07_encabe <> "" Then
'        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!d07_encabe) + 1) = DbEncerrantes.Recordset!d07_encabe
'      End If
'    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset!enc_fechamento_lts) = False Then
      If DbEncerrantes.Recordset!enc_fechamento_lts <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!enc_fechamento_lts) + 1) = DbEncerrantes.Recordset!enc_fechamento_lts
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    
    With dbDestino
      .Recordset.AddNew
      .Recordset!DataCaixa = Dia
      .Recordset!Turno = Turno
      .Recordset!CodigoPosto = CodigoPosto
      .Recordset!NomePosto = NomePosto
      .Recordset!linhaexportada = StrLinha
      .Recordset.Update
    End With
    
    DbEncerrantes.Recordset.MoveNext
  Loop
End If

If dbVendas.Recordset.RecordCount <> 0 Then
  dbVendas.Recordset.MoveLast
  dbVendas.Recordset.MoveFirst
  Do While dbVendas.Recordset.EOF = False
    lblStatus.Caption = "Exportando Vendas... " & Format((dbVendas.Recordset.AbsolutePosition / (dbVendas.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!produto_id) = False Then
      If dbVendas.Recordset!produto_id <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!produto_id) + 1) = dbVendas.Recordset!produto_id
      End If
    End If
    StrLinha = "002|" & StrTemp
    
    StrTemp = Space(6)
'    If IsNull(dbVendas.Recordset!codbom) = False Then
'      If dbVendas.Recordset!codbom <> "" Then
'        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!codbom) + 1) = dbVendas.Recordset!codbom
'      End If
'    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!qtd_venda) = False Then
      If dbVendas.Recordset!qtd_venda <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!qtd_venda) + 1) = dbVendas.Recordset!qtd_venda
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!prc_venda) = False Then
      If dbVendas.Recordset!prc_venda <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!prc_venda) + 1) = dbVendas.Recordset!prc_venda
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!val_total) = False Then
      If dbVendas.Recordset!val_total <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!val_total) + 1) = dbVendas.Recordset!val_total
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!codigo) = False Then
      If dbVendas.Recordset!pessoa_id <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!codigo) + 1) = dbVendas.Recordset!codigo
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    With dbDestino
      .Recordset.AddNew
      .Recordset!DataCaixa = Dia
      .Recordset!Turno = Turno
      .Recordset!CodigoPosto = CodigoPosto
      .Recordset!NomePosto = NomePosto
      .Recordset!linhaexportada = StrLinha
      .Recordset.Update
    End With
    
    dbVendas.Recordset.MoveNext
  Loop
End If

If dbClientes.Recordset.RecordCount <> 0 Then
  dbClientes.Recordset.MoveLast
  dbClientes.Recordset.MoveFirst
  Do While dbClientes.Recordset.EOF = False
    lblStatus.Caption = "Exportando Clientes... " & Format((dbClientes.Recordset.AbsolutePosition / (dbClientes.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    If IsNull(dbClientes.Recordset!pessoa_id) = False Then
      If dbClientes.Recordset!pessoa_id <> "" Then
        StrTemp = Space(12)
        If IsNull(dbClientes.Recordset!pessoa_id) = False Then
          If dbPessoa.Recordset.RecordCount <> 0 Then
            dbPessoa.Recordset.MoveLast
            dbPessoa.Recordset.MoveFirst
            dbPessoa.Recordset.Find "pessoa_id=" & dbClientes.Recordset!pessoa_id
            If dbPessoa.Recordset.EOF = False Then
              Mid(StrTemp, Len(StrTemp) - Len(dbPessoa.Recordset!codigo) + 1) = dbPessoa.Recordset!codigo
            End If
          End If
        End If
        StrLinha = "003|" & StrTemp
        
        StrTemp = Space(12)
        If IsNull(dbClientes.Recordset("num_nf")) = False Then
          If dbClientes.Recordset("num_nf") <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset("num_nf")) + 1) = dbClientes.Recordset("num_nf")
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(9)
        
        If IsNull(dbClientes.Recordset!placa_id) = False Then
          If dbClientes.Recordset!placa_id <> "" Then
            Mid(StrTemp, 1) = dbClientes.Recordset!placa_id
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15)
        If IsNull(dbClientes.Recordset!km) = False Then
          If dbClientes.Recordset!km <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!km) + 1) = dbClientes.Recordset!km
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(25)
        If IsNull(dbClientes.Recordset!marca) = False Then
          If dbClientes.Recordset!marca <> "" Then
            Mid(StrTemp, 1) = dbClientes.Recordset!marca
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
      Else
        StrTemp = Space(9)
        StrLinha = StrLinha & "|" & StrTemp
        StrTemp = Space(15)
        StrLinha = StrLinha & "|" & StrTemp
        StrTemp = Space(25)
        StrLinha = StrLinha & "|" & StrTemp
      End If
          
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!qtde) = False Then
        If dbClientes.Recordset!qtde <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!qtde) + 1) = dbClientes.Recordset!qtde
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!val_item) = False Then
        Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!val_item) + 1) = dbClientes.Recordset!val_item
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!produto_id) = False Then
        If dbClientes.Recordset!produto_id <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!produto_id) + 1) = dbClientes.Recordset!produto_id
        End If
      Else
        If IsNull(dbClientes.Recordset!combustivel_id) = False Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!combustivel_id) + 1) = dbClientes.Recordset!combustivel_id
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
'      If IsNull(dbClientes.Recordset!vl_desconto) = False Then
'        If dbClientes.Recordset!vl_desconto <> 0 Then
'          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_desconto) + 1) = dbClientes.Recordset!vl_desconto
'        End If
'      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!prc_unitario) = False Then
        If dbClientes.Recordset!prc_unitario <> 0 Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!prc_unitario) + 1) = dbClientes.Recordset!prc_unitario
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!val_item) = False Then
        If dbClientes.Recordset!val_item <> 0 Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!val_item) + 1) = dbClientes.Recordset!val_item
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      
      With dbDestino
        .Recordset.AddNew
        .Recordset!DataCaixa = Dia
        .Recordset!Turno = Turno
        .Recordset!CodigoPosto = CodigoPosto
        .Recordset!NomePosto = NomePosto
        .Recordset!linhaexportada = StrLinha
        .Recordset.Update
      End With
    End If
    dbClientes.Recordset.MoveNext
  Loop
End If


If dbInventario.Recordset.RecordCount <> 0 Then
  dbInventario.Recordset.MoveLast
  dbInventario.Recordset.MoveFirst
  Do While dbInventario.Recordset.EOF = False
    lblStatus.Caption = "Exportando Inventário... " & Format((dbInventario.Recordset.AbsolutePosition / (dbInventario.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    StrTemp = Space(5)
    If IsNull(dbInventario.Recordset!tanque_id) = False Then
      If dbInventario.Recordset!tanque_id <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbInventario.Recordset!tanque_id) + 1) = dbInventario.Recordset!tanque_id
      End If
    End If
    StrLinha = "004|" & StrTemp
    
    StrTemp = Space(10)
    If IsNull(dbInventario.Recordset!qtde) = False Then
      If dbInventario.Recordset!qtde <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbInventario.Recordset!qtde) + 1) = dbInventario.Recordset!qtde
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    With dbDestino
      .Recordset.AddNew
      .Recordset!DataCaixa = Dia
      .Recordset!Turno = Turno
      .Recordset!CodigoPosto = CodigoPosto
      .Recordset!NomePosto = NomePosto
      .Recordset!linhaexportada = StrLinha
      .Recordset.Update
    End With
    
    dbInventario.Recordset.MoveNext
  Loop
End If

With dbNumerarios
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      lblStatus.Caption = "Exportando Numerários... " & Format((dbNumerarios.Recordset.AbsolutePosition / (dbNumerarios.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh
      
      StrTemp = Space(15)
      If IsNull(.Recordset!cartao_id) = False Then
        If .Recordset!cartao_id <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!cartao_id) + 1) = .Recordset!cartao_id
        End If
      End If
      StrLinha = "005|" & StrTemp
      
      StrTemp = Space(15)
'      If IsNull(.Recordset!Documento) = False Then
'        If .Recordset!Documento <> "" Then
'          Mid(StrTemp, 1) = .Recordset!Documento
'        End If
'      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Total) = False Then
        If .Recordset!Total <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!Total) + 1) = .Recordset!Total
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      With dbDestino
        .Recordset.AddNew
        .Recordset!DataCaixa = Dia
        .Recordset!Turno = Turno
        .Recordset!CodigoPosto = CodigoPosto
        .Recordset!NomePosto = NomePosto
        .Recordset!linhaexportada = StrLinha
        .Recordset.Update
      End With
      .Recordset.MoveNext
    Loop
  End If
End With

With dbDespesas
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      lblStatus.Caption = "Exportando Despesas... " & Format((.Recordset.AbsolutePosition / (.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh

      StrTemp = Space(15)
      If IsNull(.Recordset!despesa_nivel3_id) = False Then
        If .Recordset!despesa_nivel3_id <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!despesa_nivel3_id) + 1) = .Recordset!despesa_nivel3_id
        End If
      End If
      StrLinha = "006|" & StrTemp

      StrTemp = Space(50)
      If IsNull(.Recordset!descricao) = False Then
        If .Recordset!descricao <> "" Then
          Mid(StrTemp, 1) = .Recordset!descricao
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp

      StrTemp = Space(5)
      'REC - Recebimento
      'PAG - Pagamento
'      If IsNull(.Recordset!tipo) = False Then
'        If .Recordset!tipo <> "" Then
'          Mid(StrTemp, 1) = .Recordset!tipo
'        End If
'      End If
      StrLinha = StrLinha & "|" & StrTemp

      StrTemp = Space(15)
      If IsNull(.Recordset!val_despesa) = False Then
        If .Recordset!val_despesa <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(-.Recordset!val_despesa) + 1) = -.Recordset!val_despesa
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp

      With dbDestino
        .Recordset.AddNew
        .Recordset!DataCaixa = Dia
        .Recordset!Turno = Turno
        .Recordset!CodigoPosto = CodigoPosto
        .Recordset!NomePosto = NomePosto
        .Recordset!linhaexportada = StrLinha
        .Recordset.Update
      End With
      .Recordset.MoveNext
    Loop
  End If
End With

lblStatus.Caption = ""
Exit Sub

TrataErro:

MsgBox Err.Description

End Sub

Private Sub cboTurno_LostFocus()
With cboTurno
  .Text = Format(.Text, "00")
End With
End Sub

Private Sub cmdExportar_Click()
If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If
If cboTurno.Text = "" Then
  MsgBox "Informe um turno inicial!"
  cboTurno.SetFocus
  Exit Sub
End If
Animation1.Visible = True
Animation1.Open App.Path & "\engrenagem.avi"
Animation1.Play
PegaEncerrantes txtDataCaixa.Value, cboTurno.Text
Animation1.Visible = False
txtDataCaixa.SetFocus
GravarConfiguracoes Caminho, Destino, txtDataCaixa.Value, cboTurno.Text, strMSDE, CodigoPosto, NomePosto
End Sub

Private Sub cmdSair_Click()
End
End Sub

Private Sub Configurar_Click()
frmConfigura.Show vbModal
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
'firebird
'Provider=MSDASQL.1;Persist Security Info=False;Extended Properties="DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; DBNAME=c:\rede\PFWIN - Cópia.GDB;"

frmSplash.Show
frmSplash.Refresh
frmSplash.SetFocus
On Error Resume Next
frmSplash.lblWarning.Caption = "Inicializando o sistema..."
frmSplash.lblWarning.Refresh

frmSplash.lblWarning.Caption = "Carregando configurações..."

PegaConfiguracoes

If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If

frmSplash.lblWarning.Refresh

Unload frmSplash
FbUsuario = "SYSDBA"
FbSenha = "masterkey"

End Sub

Private Sub txtDataCaixa_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataCaixa_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataCaixa_LostFocus()
Me.KeyPreview = True
End Sub

