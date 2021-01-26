VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmIntegrador 
   Caption         =   "Integrador 2.0"
   ClientHeight    =   3030
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4500
   Icon            =   "Integrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdProdutos 
      Caption         =   "Exporta Grupos de Produtos"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3480
      TabIndex        =   8
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton Configurar 
      Caption         =   "Configurar"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3495
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Visible         =   0   'False
      Width           =   4815
      Begin MSAdodcLib.Adodc dbDestino 
         Height          =   330
         Left            =   2280
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
      Begin MSAdodcLib.Adodc DbEncerrantes 
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
      Begin MSAdodcLib.Adodc dbVendas 
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
         Height          =   330
         Left            =   240
         Top             =   1440
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
         Connect         =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=NEWTREND"
         OLEDBString     =   "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=NEWTREND"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   "sa"
         Password        =   "masterkey"
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
      Begin MSAdodcLib.Adodc dbInventario 
         Height          =   330
         Left            =   240
         Top             =   1800
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
      Begin MSAdodcLib.Adodc dbNumerarios 
         Height          =   330
         Left            =   240
         Top             =   2160
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
      Begin MSAdodcLib.Adodc dbDespesas 
         Height          =   330
         Left            =   240
         Top             =   2520
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
      Begin MSAdodcLib.Adodc storExcluirCaixa 
         Height          =   330
         Left            =   2280
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   65535
      Left            =   2400
      Top             =   840
   End
   Begin vbskfr2.Skinner Skinner1 
      Left            =   2760
      Top             =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   855
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
      Format          =   16515073
      CurrentDate     =   39034
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   529
      _Version        =   393216
      Format          =   16515073
      CurrentDate     =   39034
   End
   Begin VB.Label Label3 
      Caption         =   "Data Final:"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   120
      Width           =   1335
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
End
Attribute VB_Name = "frmIntegrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CodigoPosto As String, NomePosto As String
Public Caminho As String, Conectar As String, Destino As String
Public WsPrincipal As Workspace, DbPrincipal As Database
Public strMSDE As String, IniciaComWindows As Boolean

Private Sub AtualizaAdo()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

On Error Resume Next
db.Open "Provider=" & strMSDE & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & Destino
If Err.Number <> 0 Then Exit Sub

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select planodeconta from caixas order by PlanoDeConta", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "ALTER TABLE Caixas Add PlanoDeConta nVarChar(20)"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  Else
    db.Execute "update caixas set PlanoDeConta='2100000000'"
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from GruposECF", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create TABLE GruposECF (Posto nVarChar(3), Grupo nvarchar(3), Descri nvarchar(50))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  End If
End If
dbTemp.Close

On Error GoTo 0
On Error Resume Next
dbTemp.Open "select *from ProdutosGrupo", db, adOpenForwardOnly, adLockReadOnly
If Err.Number <> 0 Then
  On Error GoTo 0
  On Error Resume Next
  db.Execute "Create TABLE ProdutosGrupo (Posto nVarChar(3), CodProduto bigint, Descri nvarchar(50), Grupo nvarchar(3), GrupoECF nvarchar(3))"
  If Err.Number <> 0 Then
    MsgBox "Erro " & Err.Number & " - " & Err.Description
  End If
End If
dbTemp.Close

db.Close

End Sub

Public Function GravarConfiguracoes(ByVal strOrigem As String, ByVal strDestino As String, ByVal dtDataCaixa As String, ByVal dtDataFim As String, ByVal BancoDeDados As String, ByVal strCodigoPosto As String, ByVal strNomePosto As String) As Boolean
GravarConfiguracoes = False
SaveSetting App.EXEName, "Config", "Codigo", strCodigoPosto
SaveSetting App.EXEName, "Config", "Nome", strNomePosto
SaveSetting App.EXEName, "Config", "Origem", strOrigem
SaveSetting App.EXEName, "Config", "Destino", strDestino
SaveSetting App.EXEName, "Config", "Ultimo Dia", dtDataCaixa
SaveSetting App.EXEName, "Config", "Ultimo DiaFim", dtDataFim
SaveSetting App.EXEName, "Config", "MSDE", BancoDeDados

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
txtDataCaixa.Value = CDate(GetSetting(App.EXEName, "Config", "Ultimo Dia", Date))
txtDataFim.Value = CDate(GetSetting(App.EXEName, "Config", "Ultimo DiaFim", Date))

Ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador")
If Ret = "" Then
  IniciaComWindows = False
Else
  IniciaComWindows = True
End If

End Function

Private Sub PegaCaixas()
Dim DataCaixa As Date, Turno As String
Dim strDb As String
On Error GoTo TrataErro
If LCase(Mid(Caminho, 1, 3)) = "c:\" Then Exit Sub
strDb = "A30Sigpo"
Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
With dbCaixa
  .ConnectionString = Conectar
  .RecordSource = "select *from a30caixa order by dt_movto, turno"
  .Refresh
  DataCaixa = DateAdd("m", -1, Date)
  Turno = "01"
  DataCaixa = CDate(GetSetting(App.EXEName, "Config", "Ultimo Dia", DataCaixa))
  Turno = GetSetting(App.EXEName, "Config", "Ultimo Turno", Turno)
  
  If .Recordset.RecordCount <> 0 Then
    If .Recordset.RecordCount <> 0 Then
      .Recordset.Find "dt_movto=#" & DataCaixa & "#"
      If .Recordset.EOF = True Then
        .Recordset.Filter = ""
        If .Recordset.RecordCount <> 0 Then
          dbCaixa.Recordset.MoveFirst
        End If
      Else
        .Recordset.MoveNext
      End If
      Do While .Recordset.EOF = False
        .Recordset.MoveNext
        If .Recordset!Turno = Turno Then
          Exit Do
        End If
      Loop
      Do While .Recordset.EOF = False
        DataCaixa = .Recordset!dt_movto
        Turno = .Recordset!Turno
        PegaEncerrantes DataCaixa, Turno, .Recordset!Conta
        .Recordset.MoveNext
      Loop
    End If
  End If
  'Grava qual é o procimo caixa a ser coletado
  SaveSetting App.EXEName, "Config", "Ultimo Dia", DataCaixa
  SaveSetting App.EXEName, "Config", "Ultimo Turno", Turno

End With
Exit Sub
TrataErro:
MsgBox "PegaCaixas - " & Err.Number & " - " & Err.Description
End Sub

Private Function PegaEncerrantes(ByVal Dia As Date, ByVal Turno As String, ByVal Conta As String) As Boolean
Dim strEncerrantes As String, IntArquivo As Integer
Dim StrTemp As String, StrLinha As String
Dim strDb As String, PrimeiroCupom As String, UltimoCupom As String
Dim db As New ADODB.Connection

PegaEncerrantes = False
TentaDeNovo:
On Error GoTo 0
'On Error GoTo trataErro
lblStatus.Caption = Dia & " - " & Turno & " - Verificando se existe registros já exportados..."
lblStatus.Refresh
With storExcluirCaixa
  .ConnectionString = "Provider=" & strMSDE & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & Destino
  .RecordSource = "spApagaCaixa2;1('" & Dia & "','" & Turno & "','" & CodigoPosto & "','" & Conta & "')"
  On Error Resume Next
  .Refresh
  If Err.Number = -2147217900 Then
    db.Open "Provider=" & strMSDE & ";Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & Destino
    db.Execute "CREATE PROCEDURE spApagaCaixa2 @Dia DateTime, @Turno VarChar(5), @codigoposto VarChar(3), @caixa VarChar(10) As SELECT * FROM caixas WHERE datacaixa=@dia And Turno=@turno And codigoposto=@CodigoPosto And planodeconta=@caixa DELETE FROM caixas WHERE datacaixa=@dia And Turno=@turno And codigoposto=@CodigoPosto And planodeconta=@caixa"
  End If
  If Err.Number <> 0 Then
    lblStatus.Caption = ""
    Exit Function
  End If
End With
On Error GoTo 0
lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabela de Caixas..."
lblStatus.Refresh
With dbDestino
  strDb = "Integrador"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Destino
  .ConnectionString = Conectar
  .RecordSource = "Select *from caixas where codigoposto='" & CodigoPosto & "' and datacaixa='" & Dia & "' and turno='" & Turno & "'"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabela de Bicos..."
lblStatus.Refresh
With DbEncerrantes
  strDb = "A30Sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "select a30caixa_item.bico, a30caixa_item.quantidade, a30caixa_item.vl_total, a30caixa_bicos.num_Inicial, a30caixa_bicos.Num_final from a30caixa_item, a30caixa_bicos where a30caixa_bicos.bico=a30caixa_item.bico and a30caixa_bicos.dt_movto=a30caixa_item.dt_movto and a30caixa_bicos.turno=a30caixa_item.turno and a30caixa_item.dt_movto='" & Dia & "' and a30caixa_item.turno ='" & Turno & "' and a30caixa_item.caixa='" & Conta & "' ORDER BY a30caixa_item.bico"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabelas de Vendas..."
lblStatus.Refresh
With dbVendas
  strDb = "A30Sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "Select *from a30caixa_item where dt_movto='" & Dia & "' and turno='" & Turno & "' and caixa='" & Conta & "' order by turno, DT_movto, bico"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabelas de Clientes..."
lblStatus.Refresh
With dbClientes
  strDb = "A30cupom"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "select *from a30pdv_item where dt_cupom='" & Dia & "' and turno='" & Turno & "' and cliente<>'' and caixa='" & Conta & "'"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabelas de Inventários..."
lblStatus.Refresh
With dbInventario
  strDb = "A30sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "select *from a30inventario where dt_movto='" & Dia & "' and turno='" & Turno & "'"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabelas de Numerários..."
lblStatus.Refresh
With dbNumerarios
  strDb = "A30Sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "Select *from a30CAIXA_prog where dt_movto='" & Dia & "' and turno='" & Turno & "' and caixa='" & Conta & "' order by conta"
  .Refresh
End With

lblStatus.Caption = Dia & " - " & Turno & " - Carregando Tabelas de Despesas..."
lblStatus.Refresh
With dbDespesas
  strDb = "A30Sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "Select *from a30CAIXA_recpag where dt_movto='" & Dia & "' and turno='" & Turno & "' and caixa='" & Conta & "' order by conta"
  .Refresh
End With

With dbDestino
  .Recordset.AddNew
  .Recordset!DataCaixa = Dia
  .Recordset!Turno = Turno
  .Recordset!CodigoPosto = CodigoPosto
  .Recordset!NomePosto = NomePosto
  .Recordset!linhaexportada = "000|Inicio do Caixa"
  .Recordset!planodeconta = Conta
  .Recordset.Update
End With


If DbEncerrantes.Recordset.RecordCount <> 0 Then
  DbEncerrantes.Recordset.MoveLast
  DbEncerrantes.Recordset.MoveFirst
  Do While DbEncerrantes.Recordset.EOF = False
    lblStatus.Caption = Dia & " - " & Turno & " - Exportando Bicos... " & Format((DbEncerrantes.Recordset.AbsolutePosition / (DbEncerrantes.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    StrTemp = Space(6)
    If IsNull(DbEncerrantes.Recordset!bico) = False Then
      If DbEncerrantes.Recordset!bico <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!bico) + 1) = DbEncerrantes.Recordset!bico
      End If
    End If
    StrLinha = "001|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset!num_inicial) = False Then
      If DbEncerrantes.Recordset!num_inicial <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!num_inicial) + 1) = DbEncerrantes.Recordset!num_inicial
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset!num_final) = False Then
      If DbEncerrantes.Recordset!num_final <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!num_final) + 1) = DbEncerrantes.Recordset!num_final
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset("quantidade")) = False Then
      If DbEncerrantes.Recordset("quantidade") <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset("quantidade")) + 1) = DbEncerrantes.Recordset("quantidade")
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset("vl_total")) = False Then
      If DbEncerrantes.Recordset("vl_total") <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset("vl_total")) + 1) = DbEncerrantes.Recordset("vl_total")
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
      .Recordset!planodeconta = Conta
      .Recordset.Update
    End With
    
    DbEncerrantes.Recordset.MoveNext
  Loop
End If

If dbVendas.Recordset.RecordCount <> 0 Then
  dbVendas.Recordset.MoveLast
  dbVendas.Recordset.MoveFirst
  Do While dbVendas.Recordset.EOF = False
    lblStatus.Caption = Dia & " - " & Turno & " - Exportando Vendas... " & Format((dbVendas.Recordset.AbsolutePosition / (dbVendas.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!produto) = False Then
      If dbVendas.Recordset!produto <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!produto) + 1) = dbVendas.Recordset!produto
      End If
    End If
    StrLinha = "002|" & StrTemp
    
    StrTemp = Space(6)
    If IsNull(dbVendas.Recordset!bico) = False Then
      If dbVendas.Recordset!bico <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!bico) + 1) = dbVendas.Recordset!bico
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!quantidade) = False Then
      If dbVendas.Recordset!quantidade <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!quantidade) + 1) = dbVendas.Recordset!quantidade
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!vl_venda) = False Then
      If dbVendas.Recordset!vl_venda <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!vl_venda) + 1) = dbVendas.Recordset!vl_venda
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!vl_total) = False Then
      If dbVendas.Recordset!vl_total <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!vl_total) + 1) = dbVendas.Recordset!vl_total
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!vendedor) = False Then
      If dbVendas.Recordset!vendedor <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!vendedor) + 1) = dbVendas.Recordset!vendedor
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
      .Recordset!planodeconta = Conta
      .Recordset.Update
    End With
    
    dbVendas.Recordset.MoveNext
  Loop
End If

If dbClientes.Recordset.RecordCount <> 0 Then
  dbClientes.Recordset.MoveLast
  dbClientes.Recordset.MoveFirst
  Do While dbClientes.Recordset.EOF = False
    lblStatus.Caption = Dia & " - " & Turno & " - Exportando Clientes... " & Format((dbClientes.Recordset.AbsolutePosition / (dbClientes.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    If IsNull(dbClientes.Recordset!cliente) = False Then
      If dbClientes.Recordset!cliente <> "" Then
        StrTemp = Space(12) '5
        If IsNull(dbClientes.Recordset!cliente) = False Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!cliente) + 1) = dbClientes.Recordset!cliente
        End If
        StrLinha = "003|" & StrTemp
        
        StrTemp = Space(12) '18
        If IsNull(dbClientes.Recordset!Documento) = False Then
          If dbClientes.Recordset!Documento <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!Documento) + 1) = dbClientes.Recordset!Documento
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(9) '31
        If IsNull(dbClientes.Recordset!placa) = False Then
          If dbClientes.Recordset!placa <> "" Then
            Mid(StrTemp, 1) = dbClientes.Recordset!placa
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '41
        If IsNull(dbClientes.Recordset!km) = False Then
          If dbClientes.Recordset!km <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!km) + 1) = dbClientes.Recordset!km
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(25) '57
        If IsNull(dbClientes.Recordset!carro) = False Then
          If dbClientes.Recordset!carro <> "" Then
            Mid(StrTemp, 1) = dbClientes.Recordset!carro
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '83
        If IsNull(dbClientes.Recordset!qt_produto) = False Then
          If dbClientes.Recordset!qt_produto <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!qt_produto) + 1) = dbClientes.Recordset!qt_produto
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '99
        If IsNull(dbClientes.Recordset!vl_total_item) = False Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_total_item) + 1) = dbClientes.Recordset!vl_total_item
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '115
        If IsNull(dbClientes.Recordset!produto) = False Then
          If dbClientes.Recordset!produto <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!produto) + 1) = dbClientes.Recordset!produto
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '131
        If IsNull(dbClientes.Recordset!vl_desconto) = False Then
          If dbClientes.Recordset!vl_desconto <> 0 Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_desconto) + 1) = dbClientes.Recordset!vl_desconto
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '147
        If IsNull(dbClientes.Recordset!vl_venda_c) = False Then
          If dbClientes.Recordset!vl_venda_c <> 0 Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_venda_c) + 1) = dbClientes.Recordset!vl_venda_c
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '163
        If IsNull(dbClientes.Recordset!vl_total_item) = False Then
          If dbClientes.Recordset!vl_total_c <> 0 Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_total_c) + 1) = dbClientes.Recordset!vl_total_c
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15) '179
        If IsNull(dbClientes.Recordset!vl_venda_c) = False Then
          If dbClientes.Recordset!vl_venda_c <> 0 Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!vl_venda_c) + 1) = dbClientes.Recordset!vl_venda_c
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
          .Recordset!planodeconta = Conta
          .Recordset.Update
        End With
        
      End If
    End If
    dbClientes.Recordset.MoveNext
  Loop
End If


If dbInventario.Recordset.RecordCount <> 0 Then
  dbInventario.Recordset.MoveLast
  dbInventario.Recordset.MoveFirst
  Do While dbInventario.Recordset.EOF = False
    lblStatus.Caption = Dia & " - " & Turno & " - Exportando Inventário... " & Format((dbInventario.Recordset.AbsolutePosition / (dbInventario.Recordset.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    
    StrTemp = Space(5)
    If IsNull(dbInventario.Recordset!tanque) = False Then
      If dbInventario.Recordset!tanque <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbInventario.Recordset!tanque) + 1) = dbInventario.Recordset!tanque
      End If
    End If
    StrLinha = "004|" & StrTemp
    
    StrTemp = Space(10)
    If IsNull(dbInventario.Recordset!quantidade) = False Then
      If dbInventario.Recordset!quantidade <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbInventario.Recordset!quantidade) + 1) = dbInventario.Recordset!quantidade
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
      .Recordset!planodeconta = Conta
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
      lblStatus.Caption = Dia & " - " & Turno & " - Exportando Numerários... " & Format((dbNumerarios.Recordset.AbsolutePosition / (dbNumerarios.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Conta) = False Then
        If .Recordset!Conta <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!Conta) + 1) = .Recordset!Conta
        End If
      End If
      StrLinha = "005|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Documento) = False Then
        If .Recordset!Documento <> "" Then
          Mid(StrTemp, 1) = .Recordset!Documento
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!valor) = False Then
        If .Recordset!valor <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!valor) + 1) = .Recordset!valor
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
        .Recordset!planodeconta = Conta
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
      lblStatus.Caption = Dia & " - " & Turno & " - Exportando Despesas... " & Format((.Recordset.AbsolutePosition / (.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh
      
      StrTemp = Space(15)
      If IsNull(.Recordset!Conta) = False Then
        If .Recordset!Conta <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!Conta) + 1) = .Recordset!Conta
        End If
      End If
      StrLinha = "006|" & StrTemp
      
      StrTemp = Space(50)
      If IsNull(.Recordset!complemento) = False Then
        If .Recordset!complemento <> "" Then
          Mid(StrTemp, 1) = .Recordset!complemento
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(5)
      If IsNull(.Recordset!tipo) = False Then
        If .Recordset!tipo <> "" Then
          Mid(StrTemp, 1) = .Recordset!tipo
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!valor) = False Then
        If .Recordset!valor <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!valor) + 1) = .Recordset!valor
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
        .Recordset!planodeconta = Conta
        .Recordset.Update
      End With
      .Recordset.MoveNext
    Loop
  End If
End With


With dbClientes
  strDb = "A30sigpo"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
  .ConnectionString = Conectar
  .RecordSource = "select *from a30caixa_conta where dt_movto='" & Dia & "' and turno='" & Turno & "' and caixa='" & Conta & "'"
  .Refresh
  '.Recordset.Filter = "dt_cupom=#" & Dia & "# and turno='" & Turno & "'"
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      lblStatus.Caption = Dia & " - " & Turno & " - Exportando Resultado do caixa... " & Format((dbClientes.Recordset.AbsolutePosition / (dbClientes.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh
      StrTemp = Space(15)
      If IsNull(.Recordset!Conta) = False Then
        If .Recordset!Conta <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!Conta) + 1) = .Recordset!Conta
        End If
      End If
      StrLinha = "998|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(.Recordset!movimento) = False Then
        Mid(StrTemp, 1) = .Recordset!movimento
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      With dbDestino
        .Recordset.AddNew
        .Recordset!DataCaixa = Dia
        .Recordset!Turno = Turno
        .Recordset!CodigoPosto = CodigoPosto
        .Recordset!NomePosto = NomePosto
        .Recordset!linhaexportada = StrLinha
        .Recordset!planodeconta = Conta
        .Recordset.Update
      End With
      
      .Recordset.MoveNext
    Loop
  End If
End With

If Turno = "01" Then
  With dbClientes
    strDb = "A30CUPOM"
    Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
    .ConnectionString = Conectar
    .RecordSource = "select dt_cupom, hora, turno, documento from a30pdv_item where dt_cupom between '" & DateAdd("d", -1, Dia) & "' and '" & DateAdd("d", 1, Dia) & "' order by dt_cupom, hora"
    .Refresh
    PrimeiroCupom = "0"
    UltimoCupom = "0"
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveFirst
      .Recordset.Find "dt_cupom='" & Dia & "'"
      If .Recordset.EOF = False Then
        If .Recordset!Turno = "01" Then
          Do
            .Recordset.MovePrevious
            If .Recordset.BOF = True Then Exit Do
          Loop While .Recordset!Turno = "01"
          .Recordset.MoveNext
          PrimeiroCupom = .Recordset!Documento
        Else
          Do
            .Recordset.MoveNext
            If .Recordset.EOF = True Then Exit Do
          Loop While .Recordset!Turno <> "01"
          .Recordset.MovePrevious
          PrimeiroCupom = .Recordset!Documento
        End If
      End If
      
      .Recordset.MoveFirst
      .Recordset.Find "dt_cupom='" & DateAdd("d", 1, Dia) & "'"
      If .Recordset.EOF = False Then
        If .Recordset!Turno = "01" Then
          Do
            .Recordset.MovePrevious
            If .Recordset.BOF = True Then Exit Do
          Loop While .Recordset!Turno = "01"
          .Recordset.MoveNext
          UltimoCupom = .Recordset!Documento
        Else
          Do
            .Recordset.MoveNext
            If .Recordset.EOF = True Then Exit Do
          Loop While .Recordset!Turno <> "01"
          .Recordset.MovePrevious
          UltimoCupom = .Recordset!Documento
        End If
      End If
      
      
      
    End If
    
    .RecordSource = "select dt_cupom, sum(qt_produto) as Quantidade, sum(vl_total_item) as total, grupoif from a30pdv_item where Sg_Cancelado=0 and documento between '" & PrimeiroCupom & "' and '" & UltimoCupom & "' group by dt_cupom, grupoif order by dt_cupom"
    .Refresh
    '.Recordset.Filter = "dt_cupom=#" & Dia & "# and turno='" & Turno & "'"
  End With
  
  If dbClientes.Recordset.RecordCount <> 0 Then
    dbClientes.Recordset.MoveLast
    dbClientes.Recordset.MoveFirst
    Do While dbClientes.Recordset.EOF = False
      lblStatus.Caption = Dia & " - " & Turno & " - Exportando Cupons Fiscais... " & Format((dbClientes.Recordset.AbsolutePosition / (dbClientes.Recordset.RecordCount + 1)) * 100, "###") & "%"
      lblStatus.Refresh
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!dt_cupom) = False Then
        If dbClientes.Recordset!dt_cupom <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!dt_cupom) + 1) = dbClientes.Recordset!dt_cupom
        End If
      End If
      StrLinha = "007|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!quantidade) = False Then
        Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!quantidade) + 1) = dbClientes.Recordset!quantidade
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!Total) = False Then
        Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!Total) + 1) = dbClientes.Recordset!Total
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!grupoif) = False Then
        If dbClientes.Recordset!grupoif <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!grupoif) + 1) = dbClientes.Recordset!grupoif
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
        .Recordset!planodeconta = Conta
        .Recordset.Update
      End With
      
      dbClientes.Recordset.MoveNext
    Loop
  End If
End If

With dbDestino
  .Recordset.AddNew
  .Recordset!DataCaixa = Dia
  .Recordset!Turno = Turno
  .Recordset!CodigoPosto = CodigoPosto
  .Recordset!NomePosto = NomePosto
  .Recordset!linhaexportada = "999|Fim do Caixa"
  .Recordset!planodeconta = Conta
  .Recordset.Update
End With

lblStatus.Caption = ""
PegaEncerrantes = True

End Function

Private Sub cmdExportar_Click()
Dim db As New ADODB.Connection
Dim dbCaixas As New ADODB.Recordset

If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If

Animation1.Visible = True
Animation1.Open App.Path & "\engrenagem.avi"
Animation1.Play

strDb = "A30Sigpo"
Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Password=masterkey;Initial Catalog=" & strDb & " ;Data Source=" & Caminho

db.Open Conectar
dbCaixas.CursorLocation = adUseClient
dbCaixas.Open "Select *from a30caixa where dt_movto between '" & txtDataCaixa.Value & "' and '" & txtDataFim.Value & "'", db, adOpenKeyset, adLockOptimistic

If dbCaixas.RecordCount <> 0 Then
  Do While dbCaixas.EOF = False
    Turno = dbCaixas!Turno
    If PegaEncerrantes(dbCaixas!dt_movto, Turno, dbCaixas!caixa) = False Then
      Resposta = MsgBox("Houve um erro durante a exportação do caixa atual! Deseja tentar novamente?", vbYesNo)
      If Resposta = vbNo Then Exit Sub
    Else
      dbCaixas.MoveNext
    End If
    
  Loop
End If

Animation1.Visible = False
txtDataCaixa.SetFocus
GravarConfiguracoes Caminho, Destino, txtDataCaixa.Value, txtDataFim.Value, strMSDE, CodigoPosto, NomePosto
End Sub

Private Sub cmdProdutos_Click()
Dim strDb As String, Conectar As String
Dim db As New ADODB.Connection
Dim db2 As New ADODB.Connection
Dim dbProdutos As New ADODB.Recordset
Dim dbGrupos As New ADODB.Recordset



On Error GoTo 0
On Error Resume Next
With dbDestino
  strDb = "Integrador"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Password=masterkey;Initial Catalog=" & strDb & " ;Data Source=" & Destino
  .ConnectionString = Conectar
  .RecordSource = "Select *from GruposECF where posto='" & CodigoPosto & "'"
  .Refresh
  If Err.Number <> 0 Then
    AtualizaAdo
    .Refresh
  End If
End With

On Error GoTo 0
strDb = "A30Sigpo"
db.Open "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Password=masterkey;Initial Catalog=" & strDb & " ;Data Source=" & Caminho
dbGrupos.CursorLocation = adUseClient
dbGrupos.Open "Select *from a30grupo", db, adOpenForwardOnly, adLockReadOnly

strDb = "Integrador"
db2.Open "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Password=masterkey;Initial Catalog=" & strDb & " ;Data Source=" & Destino

db2.Execute "delete from gruposecf where posto='" & CodigoPosto & "'"
db2.Execute "delete from produtosgrupo where posto='" & CodigoPosto & "'"

If dbGrupos.RecordCount <> 0 Then
  Do While dbGrupos.EOF = False
    lblStatus.Caption = "Exportando Grupos... " & Format((dbGrupos.AbsolutePosition / (dbGrupos.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    db2.Execute "insert into gruposecf (posto,grupo,descri) values ('" & CodigoPosto & "','" & dbGrupos!grupo & "','" & dbGrupos!descricao & "')"
    dbGrupos.MoveNext
  Loop
End If
dbGrupos.Close

dbProdutos.CursorLocation = adUseClient
dbProdutos.Open "SELECT produto, descricao, grupo, grupoif FROM A30PRODUTO", db, adOpenForwardOnly, adLockReadOnly
If dbProdutos.RecordCount <> 0 Then
  Do While dbProdutos.EOF = False
    lblStatus.Caption = "Exportando Produtos... " & Format((dbProdutos.AbsolutePosition / (dbProdutos.RecordCount + 1)) * 100, "###") & "%"
    lblStatus.Refresh
    db2.Execute "insert into produtosgrupo (Posto,CodProduto,Descri,Grupo,GrupoECF) values ('" & CodigoPosto & "'," & dbProdutos!produto & ",'" & dbProdutos!descricao & "','" & dbProdutos!grupo & "','" & dbProdutos!grupoif & "')"
    dbProdutos.MoveNext
  Loop
End If
dbProdutos.Close
db.Close
db2.Close
lblStatus.Caption = "Exportação completa."
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
'Provider=SQLNCLI.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=newtrend1
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=NEWTREND
frmSplash.Show
frmSplash.Refresh
frmSplash.SetFocus
On Error Resume Next
frmSplash.lblWarning.Caption = "Inicializando o sistema..."
frmSplash.lblWarning.Refresh

Me.Caption = "Integrador Versão " & App.Major & "." & App.Minor & "." & App.Revision

frmSplash.lblWarning.Caption = "Carregando configurações..."

PegaConfiguracoes

If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If

PegaConfiguracoes

AtualizaAdo

frmSplash.lblWarning.Refresh


Unload frmSplash
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

Private Sub txtDataFim_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataFim_LostFocus()
Me.KeyPreview = True
End Sub

