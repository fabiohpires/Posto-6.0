VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{2A51FC74-DB07-4C60-B0BC-71F1A20E713D}#1.0#0"; "vbskfr2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmIntegrador 
   Caption         =   "Integrador p/ Posto Fácil"
   ClientHeight    =   3015
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4485
   Icon            =   "Integrador.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3375
      Left            =   720
      TabIndex        =   5
      Top             =   2520
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
      Format          =   16646145
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
Public Caminho As String, Conectar As String, Destino As String
Public WsPrincipal As Workspace, DbPrincipal As Database
Public strMSDE As String, IniciaComWindows As Boolean
Public strDBase As String

Public Function GravarConfiguracoes(ByVal strOrigem As String, ByVal strDestino As String, ByVal dtDataCaixa As String, ByVal strTurno As String, ByVal BancoDeDados As String, ByVal strCodigoPosto As String, ByVal strNomePosto As String) As Boolean
GravarConfiguracoes = False
SaveSetting App.EXEName, "Config", "Codigo", strCodigoPosto
SaveSetting App.EXEName, "Config", "Nome", strNomePosto
SaveSetting App.EXEName, "Config", "Origem", strOrigem
SaveSetting App.EXEName, "Config", "Destino", strDestino
SaveSetting App.EXEName, "Config", "Ultimo Dia", dtDataCaixa
SaveSetting App.EXEName, "Config", "Ultimo Turno", strTurno
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
cboTurno.Text = GetSetting(App.EXEName, "Config", "Ultimo Turno", "01")
txtDataCaixa.Value = CDate(GetSetting(App.EXEName, "Config", "Ultimo Dia", Date))

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
With DbEncerrantes
  .DatabaseName = Caminho
  .RecordSource = "Select * from CF00D007 where d07_datmov=#" & DataInglesa(Dia) & "# and d07_fecatu='" & CInt(Turno) & "'"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Vendas..."
lblStatus.Refresh
With dbVendas
  .DatabaseName = Caminho
  .RecordSource = "Select cf00d008.*, cf00d011.* from cf00d008, cf00d011 where cf00d008.numref=cf00d011.numref and datz=#" & DataInglesa(Dia) & "# and numfec='" & CInt(Turno) & "' and tippro<>'' order by datmov, numfec, codpro"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Clientes..."
lblStatus.Refresh
With dbClientes
  .DatabaseName = Caminho
  .RecordSource = "Select cf00d008.*, cf00d011.* from cf00d008, cf00d011 where cf00d008.numref=cf00d011.numref and datz=#" & DataInglesa(Dia) & "# and numfec='" & CInt(Turno) & "' and codcli<>'' order by datmov, numfec"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Inventários..."
lblStatus.Refresh
With dbInventario
  .DatabaseName = Caminho
  .RecordSource = "select *from cf00d014 where datmed=#" & DataInglesa(Dia) & "# and numfec='" & CInt(Turno) & "'"
  .Refresh
End With

lblStatus.Caption = "Carregando Tabelas de Numerários..."
lblStatus.Refresh
With dbNumerarios
  .DatabaseName = Caminho
  .RecordSource = "Select datz, numfec, codcar, sum(valpag) as total from cf00d008 where datz=#" & DataInglesa(Dia) & "# and numfec='" & CInt(Turno) & "' and conpag='R' group by codcar, datz, numfec order by codcar"
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
    If IsNull(DbEncerrantes.Recordset!d07_codbom) = False Then
      If DbEncerrantes.Recordset!d07_codbom <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(Mid(DbEncerrantes.Recordset!d07_codbom, 2)) + 1) = Mid(DbEncerrantes.Recordset!d07_codbom, 2)
      End If
    End If
    StrLinha = "001|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset!d07_encabe) = False Then
      If DbEncerrantes.Recordset!d07_encabe <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!d07_encabe) + 1) = DbEncerrantes.Recordset!d07_encabe
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(16)
    If IsNull(DbEncerrantes.Recordset!d07_encfec) = False Then
      If DbEncerrantes.Recordset!d07_encfec <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(DbEncerrantes.Recordset!d07_encfec) + 1) = DbEncerrantes.Recordset!d07_encfec
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
    If IsNull(dbVendas.Recordset!codpro) = False Then
      If dbVendas.Recordset!codpro <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!codpro) + 1) = dbVendas.Recordset!codpro
      End If
    End If
    StrLinha = "002|" & StrTemp
    
    StrTemp = Space(6)
    If IsNull(dbVendas.Recordset!codbom) = False Then
      If dbVendas.Recordset!codbom <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!codbom) + 1) = dbVendas.Recordset!codbom
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!qtde) = False Then
      If dbVendas.Recordset!qtde <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!qtde) + 1) = dbVendas.Recordset!qtde
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!prcuni) = False Then
      If dbVendas.Recordset!prcuni <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!prcuni) + 1) = dbVendas.Recordset!prcuni
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!valpro) = False Then
      If dbVendas.Recordset!valpro <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!valpro) + 1) = dbVendas.Recordset!valpro
      End If
    End If
    StrLinha = StrLinha & "|" & StrTemp
    
    StrTemp = Space(12)
    If IsNull(dbVendas.Recordset!codfun) = False Then
      If dbVendas.Recordset!codfun <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbVendas.Recordset!codfun) + 1) = dbVendas.Recordset!codfun
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
    
    If IsNull(dbClientes.Recordset!codcli) = False Then
      If dbClientes.Recordset!codcli <> "" Then
        StrTemp = Space(12)
        If IsNull(dbClientes.Recordset!codcli) = False Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!codcli) + 1) = dbClientes.Recordset!codcli
        End If
        StrLinha = "003|" & StrTemp
        
        StrTemp = Space(12)
        If IsNull(dbClientes.Recordset("cf00d008.numref")) = False Then
          If dbClientes.Recordset("cf00d008.numref") <> "" Then
            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset("cf00d008.numref")) + 1) = dbClientes.Recordset("cf00d008.numref")
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(9)
        
        If IsNull(dbClientes.Recordset!placa) = False Then
          If dbClientes.Recordset!placa <> "" Then
            Mid(StrTemp, 1) = dbClientes.Recordset!placa
          End If
        End If
        StrLinha = StrLinha & "|" & StrTemp
        
        StrTemp = Space(15)
'        If IsNull(dbClientes.Recordset!km) = False Then
'          If dbClientes.Recordset!km <> "" Then
'            Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!km) + 1) = dbClientes.Recordset!km
'          End If
'        End If
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
      If IsNull(dbClientes.Recordset!valpro) = False Then
        Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!valpro) + 1) = dbClientes.Recordset!valpro
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!codpro) = False Then
        If dbClientes.Recordset!codpro <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!codpro) + 1) = dbClientes.Recordset!codpro
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
      If IsNull(dbClientes.Recordset!valpro) = False Then
        If dbClientes.Recordset!valpro <> 0 Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!valpro) + 1) = dbClientes.Recordset!valpro
        End If
      End If
      StrLinha = StrLinha & "|" & StrTemp
      
      StrTemp = Space(15)
      If IsNull(dbClientes.Recordset!valpro) = False Then
        If dbClientes.Recordset!valpro <> 0 Then
          Mid(StrTemp, Len(StrTemp) - Len(dbClientes.Recordset!valpro) + 1) = dbClientes.Recordset!valpro
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
    If IsNull(dbInventario.Recordset!codtan) = False Then
      If dbInventario.Recordset!codtan <> "" Then
        Mid(StrTemp, Len(StrTemp) - Len(dbInventario.Recordset!codtan) + 1) = dbInventario.Recordset!codtan
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
      If IsNull(.Recordset!codcar) = False Then
        If .Recordset!codcar <> "" Then
          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!codcar) + 1) = .Recordset!codcar
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

'With dbDespesas
'  If .Recordset.RecordCount <> 0 Then
'    .Recordset.MoveLast
'    .Recordset.MoveFirst
'    Do While .Recordset.EOF = False
'      lblStatus.Caption = "Exportando Despesas... " & Format((.Recordset.AbsolutePosition / (.Recordset.RecordCount + 1)) * 100, "###") & "%"
'      lblStatus.Refresh
'
'      StrTemp = Space(15)
'      If IsNull(.Recordset!conta) = False Then
'        If .Recordset!conta <> "" Then
'          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!conta) + 1) = .Recordset!conta
'        End If
'      End If
'      StrLinha = "006|" & StrTemp
'
'      StrTemp = Space(50)
'      If IsNull(.Recordset!complemento) = False Then
'        If .Recordset!complemento <> "" Then
'          Mid(StrTemp, 1) = .Recordset!complemento
'        End If
'      End If
'      StrLinha = StrLinha & "|" & StrTemp
'
'      StrTemp = Space(5)
'      If IsNull(.Recordset!tipo) = False Then
'        If .Recordset!tipo <> "" Then
'          Mid(StrTemp, 1) = .Recordset!tipo
'        End If
'      End If
'      StrLinha = StrLinha & "|" & StrTemp
'
'      StrTemp = Space(15)
'      If IsNull(.Recordset!valor) = False Then
'        If .Recordset!valor <> "" Then
'          Mid(StrTemp, Len(StrTemp) - Len(.Recordset!valor) + 1) = .Recordset!valor
'        End If
'      End If
'      StrLinha = StrLinha & "|" & StrTemp
'
'      With dbDestino
'        .Recordset.AddNew
'        .Recordset!DataCaixa = Dia
'        .Recordset!Turno = Turno
'        .Recordset!CodigoPosto = CodigoPosto
'        .Recordset!NomePosto = NomePosto
'        .Recordset!linhaexportada = StrLinha
'        .Recordset.Update
'      End With
'      .Recordset.MoveNext
'    Loop
'  End If
'End With

lblStatus.Caption = ""
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

