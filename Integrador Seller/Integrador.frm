VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIntegrador 
   Caption         =   "Integrador p/ Posto Fácil"
   ClientHeight    =   3045
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4485
   Icon            =   "Integrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin vbskfr2.Skinner Skinner1 
      Left            =   1800
      Top             =   120
      _ExtentX        =   1270
      _ExtentY        =   1270
      SysDisableSkinCaption=   "&Disable Skin"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Arquivo Texto|*.txt"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Visible         =   0   'False
      Width           =   3015
      Begin VB.Data dbNumerarios 
         Caption         =   "dbNumerarios"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2880
         Width           =   2775
      End
      Begin VB.Data dbInventario 
         Caption         =   "dbInventario"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Data dbClientes 
         Caption         =   "dbClientes"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Data dbVendas 
         Caption         =   "dbVendas"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1800
         Width           =   2775
      End
      Begin VB.Data DbEncerrantes 
         Caption         =   "DbEncerrantes"
         Connect         =   "dBASE III;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Width           =   2775
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
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Configurar 
      Caption         =   "Configurar"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
   End
   Begin MSComCtl2.Animation Animation1 
      Height          =   1215
      Left            =   120
      TabIndex        =   2
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
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   4215
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
Public strDBase As String

Public Function GravarConfiguracoes(ByVal strOrigem As String, ByVal strDestino As String, ByVal BancoDeDados As String, ByVal strCodigoPosto As String, ByVal strNomePosto As String) As Boolean
GravarConfiguracoes = False
SaveSetting App.EXEName, "Config", "Codigo", strCodigoPosto
SaveSetting App.EXEName, "Config", "Nome", strNomePosto
SaveSetting App.EXEName, "Config", "Origem", strOrigem
SaveSetting App.EXEName, "Config", "Destino", strDestino
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

Ret = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Integrador")
If Ret = "" Then
  IniciaComWindows = False
Else
  IniciaComWindows = True
End If

End Function

Private Function ImportaRegistros() As Boolean
Dim A As Integer, StrTemp As String, Dia As Date, Turno As String
Dim Contador As Double, StrLinha As String
Dim CodigoCaixa As Double, Codigocupom As Double
Dim Db As New ADODB.Connection
Dim dbCaixas As New ADODB.Recordset
Dim DbCupons As New ADODB.Recordset
Dim DbCuponsDetalhe As New ADODB.Recordset
Dim dbCuponsFormas As New ADODB.Recordset

On Error GoTo TrataErro
  With CommonDialog1
    .FileName = Caminho
    .ShowOpen
    Caminho = .FileName
  End With
End With
If Dir(Caminho) = "" Then
  MsgBox "Arquivo de exportação não localizado!"
  ImportaRegistros = False
  Exit Function
End If
TentaDeNovo:
A = FreeFile()
Contador = 1
Open Caminho For Input As #A

lblStatus.Caption = "Carregando Tabela de Caixas..."
lblStatus.Refresh
With dbDestino
  StrDb = "Integrador"
  Conectar = "Provider=" & strMSDE & ";Persist Security Info=False;User ID=sa;Initial Catalog=" & StrDb & " ;Data Source=" & Destino
  .ConnectionString = Conectar
  .RecordSource = "Select *from caixas where codigoposto='" & CodigoPosto & "' and datacaixa='" & Date & "'"
  .Refresh
End With
On Error GoTo 0

Db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Caixas.mdb"
dbCaixas.CursorLocation = adUseClient
dbCaixas.Open "Select *from Caixas", Db, adOpenKeyset, adLockOptimistic
DbCupons.CursorLocation = adUseClient
DbCupons.Open "Select *from Cupons", Db, adOpenKeyset, adLockOptimistic
DbCuponsDetalhe.CursorLocation = adUseClient
DbCuponsDetalhe.Open "Select *from CuponsDetalhe", Db, adOpenKeyset, adLockOptimistic
dbCuponsFormas.CursorLocation = adUseClient
dbCuponsFormas.Open "Select *from CuponsFormas", Db, adOpenKeyset, adLockOptimistic

Procimo:
Do While EOF(A) = False
  Line Input #A, StrTemp
  Select Case Mid(StrTemp, 1, 1)
    Case "N"
      'N#1#20090406#14:58:26#20090406#22:48:07
      Turno = Mid(StrTemp, 3, 1)
      Dia = CDate(Mid(StrTemp, 11, 2) & "/" & Mid(StrTemp, 9, 2) & "/" & Mid(StrTemp, 5, 4))
      StrTemp = "000|" & Format(Dia, "dd/mm/yyyy") & "|" & Format(Turno, "00")
      If dbCaixas.RecordCount = 0 Then
        dbCaixas.AddNew
      Else
        dbCaixas.MoveFirst
        dbCaixas.Find "datacaixa=#" & Dia & "# and turno='" & Turno & "'"
        If dbCaixas.EOF = True Then
          dbCaixas.AddNew
        Else
          CodigoCaixa = dbCaixas!CodigoCaixa
          Db.Execute "delete cupons where codigocaixa=" & CodigoCaixa
          Db.Execute "delete cuponsdetalhe where codigocaixa=" & CodigoCaixa
          Db.Execute "delete cuponsformas where codigocaixa=" & CodigoCaixa
        End If
      End If
      dbCaixas!CodigoPosto = CodigoPosto
      dbCaixas!NomePosto = NomePosto
      dbCaixas!datacaixa = Dia
      dbCaixas!Turno = Turno
      dbCaixas.Update
      CodigoCaixa = dbCaixas!CodigoCaixa
    Case "C"
      
      DbCupons.AddNew
      'C#0001#004695#20090406#15:00:35#000000190#000000000#000000000#F#00000000000000#CONSUMIDOR FINAL#1###CONSUMIDOR FINAL#
      DbCupons!CodigoCaixa = CodigoCaixa
      DbCupons!datacupom = CDate(Mid(StrTemp, 19, 2) & "/" & Mid(StrTemp, 21, 2) & "/" & Mid(StrTemp, 15, 4))
      DbCupons!numerocupom = Mid(StrTemp, 8, 6)
      DbCupons!valortotal = CCur(Mid(StrTemp, 33, 9)) / 100
      DbCupons!valordesconto = CCur(Mid(StrTemp, 43, 9)) / 100
      DbCupons!valortroco = CCur(Mid(StrTemp, 53, 9)) / 100
      If rigth(StrTemp, 1) = "C" Then
        DbCupons!cancelado = True
      Else
        DbCupons!cancelado = False
      End If
      DbCupons.Update
      Codigocupom = DbCupons!Codigocupom
    Case "I"
      'I#7896019611749#CHOC LACTA LANCY AVELA 30G#00000001000#000001900#000000190#000000000#000000000#000000190#60#00#000000000#000000000#0001#0#00000##00011#un#00002#Loja#000000087#000000086#000000070
      DbCuponsDetalhe.AddNew
      DbCuponsDetalhe!CodigoCaixa = CodigoCaixa
      DbCuponsDetalhe!Codigocupom = Codigocupom
      DbCuponsDetalhe!codigoproduto = Mid(StrTemp, 3, 13)
      A = InStr(17, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!produto = Mid(StrTemp, 17, A)
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!quantidade = CDbl(Mid(StrTemp, B, A)) / 1000
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!valorunitario = CDbl(Mid(StrTemp, B, A)) / 1000
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!valorbruto = CDbl(Mid(StrTemp, B, A)) / 100
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!valordescontoitem = CDbl(Mid(StrTemp, B, A)) / 100
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!valordescontorateadoitem = CDbl(Mid(StrTemp, B, A)) / 100
      End If
      B = A + 1
      A = InStr(B, "#", StrTemp)
      If A <> 0 Then
        DbCuponsDetalhe!valortotalitem = CDbl(Mid(StrTemp, B, A)) / 100
      End If
      If Right(StrTemp, 1) = "C" Then
        DbCuponsDetalhe!cancelado = True
      Else
        DbCuponsDetalhe!cancelado = False
      End If
      DbCuponsDetalhe.Update
    Case "P"
      
  End Select
  
  
  '**********************************************************************************
  '**********************************************************************************
gravaRegistro:
  With dbDestino
    .Recordset.AddNew
    .Recordset!datacaixa = Dia
    .Recordset!Turno = Turno
    .Recordset!CodigoPosto = CodigoPosto
    .Recordset!NomePosto = NomePosto
    .Recordset!linhaexportada = StrTemp
    .Recordset.Update
  End With
  Contador = Contador + 1
Loop

Close #A
ImportaRegistros = True
lblStatus.Caption = ""
TrataErro:
End Function

Private Sub cmdExportar_Click()
If CodigoPosto = "" Then
  MsgBox "É preciso configurar o código e o nome do posto!"
  Call Configurar_Click
End If
Animation1.Visible = True
Animation1.Open App.Path & "\engrenagem.avi"
Animation1.Play
'*******************************************************************************************
'*******************************************************************************************
ImportaRegistros

Animation1.Visible = False
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

If Dir(App.Path & "\Caixas.mdb") = "" Then
  CriaMDB
End If

'Provider=SQLNCLI.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=newtrend1
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=A30Sigpo;Data Source=NEWTREND
'Provider = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=Arquivos do dBASE;Initial Catalog=c:\backuppista"
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

End Sub
