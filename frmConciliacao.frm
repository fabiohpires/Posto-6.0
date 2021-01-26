VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConciliaCustodia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conciliação Bancária"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmConciliacao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbMovimentacao 
      Height          =   330
      Left            =   1920
      Top             =   4680
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from movimentacao"
      Caption         =   "dbMovimentacao"
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
      Left            =   7200
      Picture         =   "frmConciliacao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbDespesasLanc 
      Height          =   330
      Left            =   1920
      Top             =   4320
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from DespesasLanc"
      Caption         =   "dbDespesasLanc"
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
   Begin MSAdodcLib.Adodc dbTemp 
      Height          =   330
      Left            =   1920
      Top             =   3960
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from CompensaPendente"
      Caption         =   "dbTemp"
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
      Caption         =   "Sair"
      Height          =   375
      Left            =   7080
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin MSAdodcLib.Adodc dbPendencias 
      Height          =   330
      Left            =   1920
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from CompensaPendente"
      Caption         =   "dbPendencias"
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
   Begin VB.CommandButton cmdRecebimentos 
      Caption         =   "Recebimento"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdPagamentos 
      Caption         =   "Pagamentos"
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc dbConciliaPendente 
      Height          =   330
      Left            =   1920
      Top             =   3240
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from CompensaPendente"
      Caption         =   "dbConciliaPendente"
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
      Left            =   1920
      Top             =   2160
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
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
      Left            =   1920
      Top             =   2520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from despesatipo order by descri"
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
   Begin MSAdodcLib.Adodc dbConcilia 
      Height          =   330
      Left            =   1920
      Top             =   2880
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=posto.mdb;Mode=ReadWrite;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from concilia where codigoconta=0"
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
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin MSDataListLib.DataCombo cboContas 
      Bindings        =   "frmConciliacao.frx":0EC4
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cboDespesa 
      Bindings        =   "frmConciliacao.frx":0EDB
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmConciliacao.frx":0EF7
      Height          =   4335
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7646
      _Version        =   393216
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
      Caption         =   "Extrato da Conta"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Data"
         Caption         =   "Data"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
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
         DataField       =   "NrDocumento"
         Caption         =   "Documento"
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
      BeginProperty Column03 
         DataField       =   "Valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """ ""#.##0,00"
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
            ColumnWidth     =   780,095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3360,189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1574,929
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1454,74
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtData 
      Height          =   315
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   24641537
      CurrentDate     =   37257
   End
   Begin MSComCtl2.DTPicker txtDataLanc 
      Height          =   315
      Left            =   4800
      TabIndex        =   11
      Top             =   960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Format          =   24641537
      CurrentDate     =   37257
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "A partir de:"
      Height          =   195
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   765
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   3360
      TabIndex        =   16
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qtd.:"
      Height          =   195
      Left            =   2520
      TabIndex        =   8
      Top             =   720
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Despesa:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Conta:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmConciliaCustodia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Valor As Currency, DataBaixa As Date, Codigo As Double
Public Descri As String, NrDocumento As String

Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String

StrTemp = "Extrato de Conta"
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.CurrentY = 0
Printer.Print StrTemp

StrTemp = NomePosto
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10


StrTemp = "Data: " & Format(Dia, "long Date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Conta: " & cboContas.Text
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.Print ""

StrTemp = "Data"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Documento"
Printer.CurrentX = 100
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Public Sub TiraSaldo(ByVal DataInicial As String)
  Dim Ws As Workspace, Db As Database, dbSaldo As Recordset
  Dim Saldo As Currency, DiaAnterior As String
  
  Screen.MousePointer = vbHourglass
  
  Set Ws = DBEngine.Workspaces(0)
  Set Db = Ws.OpenDatabase(Caminho, , , Conectar)
  
  DiaAnterior = Str(DateAdd("d", -1, CDate(DataInicial)))
  
  Db.Execute "delete *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and tipo='Saldo' and data>=#" & DataInglesa(DataInicial) & "#"
  
  Set dbSaldo = Db.OpenRecordset("select *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and tipo='Saldo' and data<=#" & DataInglesa(DiaAnterior) & "# order by Data")
  If dbSaldo.RecordCount <> 0 Then
    dbSaldo.MoveLast
    Saldo = dbSaldo!Valor
  Else
    Saldo = 0
  End If
  With dbTemp
    .RecordSource = "select sum (valor) as Total, Data from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " and data>=#" & DataInglesa(DataInicial) & "#  group by data order by data"
    .Refresh
    .Refresh
    If .Recordset.RecordCount = 0 Then Exit Sub
    Do While .Recordset.EOF = False
      Saldo = Saldo + .Recordset!Total
      dbConcilia.Recordset.AddNew
      dbConcilia.Recordset!codigoconta = dbContas.Recordset!codigoconta
      dbConcilia.Recordset!Data = .Recordset!Data
      dbConcilia.Recordset!tipo = "Saldo"
      dbConcilia.Recordset!Codigo = 0
      dbConcilia.Recordset!Descri = "Saldo"
      dbConcilia.Recordset!NrDocumento = "999999999"
      dbConcilia.Recordset!Valor = Saldo
      dbConcilia.Recordset.Update
      .Recordset.MoveNext
    Loop
  End With
  dbConcilia.Refresh
  dbConcilia.Refresh
  On Error Resume Next
  dbConcilia.Recordset.MoveLast
  
  Screen.MousePointer = vbDefault
End Sub

Private Sub cboContas_LostFocus()
With dbContas
  .Refresh
  If cboContas.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboContas.Text & "'"
  If .Recordset.EOF = False Then
    cboContas.Text = .Recordset!Descri
    With dbConcilia
      .RecordSource = "select *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " order by data, codigoconciliaconta"
      .Refresh
      .Refresh
      .Refresh
      If .Recordset.RecordCount = 0 Then
        .Recordset.AddNew
        .Recordset!codigoconta = dbContas.Recordset!codigoconta
        .Recordset!Data = CDate("01/01/2002")
        .Recordset!tipo = "Saldo Inicial"
        .Recordset!Codigo = 0
        .Recordset!Descri = "Saldo Inicial"
        .Recordset!NrDocumento = "999999999"
        .Recordset!Valor = 0
        .Recordset.Update
      End If
      If IsNull(txtData.Value) = True Then
        TiraSaldo "01/01/2002"
      Else
        TiraSaldo Trim(Str(txtData.Value))
      End If
      .Recordset.MoveLast
    End With
  End If
End With
Call txtData_LostFocus
End Sub

Private Sub cboDespesa_LostFocus()
With dbDespesaTipo
  .Refresh
  If cboDespesa.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboDespesa.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  If IsNumeric(txtValor.Text) = False Then
    txtValor.Text = "1"
  End If
  lblTotal.Caption = Format((CDbl(txtValor.Text) * .Recordset!Valor), "currency")
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date
With dbConcilia
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Largura = 190
  Dia = Now
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      
      Printer.NewPage
      Cabeca Largura, Dia
      
    End If
    If .Recordset!NrDocumento = "999999999" Then
      Printer.FontBold = True
    Else
      Printer.FontBold = False
    End If
    StrTemp = .Recordset!Data
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!NrDocumento
    Printer.CurrentX = 100
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Valor, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.EndDoc
End With
naoImprime:

End Sub

Private Sub cmdIncluir_Click()
If cboContas.Text <> dbContas.Recordset!Descri Then
  MsgBox "Conta inválida! Selecione novamente."
  cboContas.SetFocus
  Exit Sub
End If
If cboDespesa.Text <> dbDespesaTipo.Recordset!Descri Then
  MsgBox "Despesa inválida! Selecione novamente."
  cboDespesa.SetFocus
  Exit Sub
End If
If IsNumeric(txtValor.Text) = False Then
  MsgBox "Valor inválido! Informe um valor correto!"
  txtValor.SetFocus
  Exit Sub
End If
If txtDataLanc.Value < DateAdd("d", -15, Date) Then
  MsgBox "Data de lançamento muito antiga!"
  Permissao = False
  frmPermissao.Show vbModal
  If Permissao = False Then
    txtDataLanc.SetFocus
    Exit Sub
  End If
End If
With dbConcilia
  .Recordset.AddNew
  .Recordset!codigoconta = dbContas.Recordset!codigoconta
  .Recordset!Data = txtDataLanc.Value
  .Recordset!tipo = "Despesa"
  .Recordset!Codigo = 0
  .Recordset!Descri = dbDespesaTipo.Recordset!Descri
  .Recordset!NrDocumento = "888888888"
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset.Update
End With
With dbContas
  .Refresh
  .Recordset.Find "descri='" & cboContas.Text & "'"
  .Recordset!Saldo = .Recordset!Saldo + CCur(lblTotal.Caption)
  .Recordset.Update
  .Refresh
  .Refresh
  .Recordset.Find "descri='" & cboContas.Text & "'"
End With
With dbDespesasLanc
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!origem = "Conciliação"
  .Recordset!Data = Date
  .Recordset!hora = Now
  .Recordset!Vencimento = txtData.Value
  .Recordset!CodigoDespesa = 0
  .Recordset!Descri = cboDespesa.Text
  .Recordset!obs = "Despesa Bancária"
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset!compensado = True
  .Recordset!fechamentodiario = True
  .Recordset.Update
  .Refresh
  .Recordset.MoveLast
End With
With dbMovimentacao
  .Recordset.AddNew
  .Recordset!Data = Now
  .Recordset!tipo = "Conciliação"
  .Recordset!codigoconta = dbContas.Recordset!codigoconta
  .Recordset!conta = dbContas.Recordset!Descri
  .Recordset!Descri = cboDespesa.Text
  .Recordset!Valor = CCur(lblTotal.Caption)
  .Recordset!Saldo = dbContas.Recordset!Saldo
  .Recordset.Update
  .Refresh
  .Refresh
End With
TiraSaldo Str(txtDataLanc.Value)
dbConcilia.Refresh
dbConcilia.Refresh
cboDespesa.Text = ""
txtValor.Text = ""
cboDespesa.SetFocus
If dbConcilia.Recordset.RecordCount <> 0 Then
  dbConcilia.Recordset.Find "data=#" & DataInglesa(Trim(Str(txtDataLanc.Value))) & "#"
End If
End Sub

Private Sub cmdPagamentos_Click()

If dbContas.Recordset.RecordCount = 0 Then Exit Sub
If dbContas.Recordset.EOF = True Then Exit Sub

Load frmConciliaPagamentos
With frmConciliaPagamentos
  With .dbPendencias
    .RecordSource = "select *from compensapendente where codigoconta=" & dbContas.Recordset!codigoconta & " and conciliado=0 order by data"
    .Refresh
    .Refresh
    .Refresh
  End With
End With
frmConciliaPagamentos.Show vbModal


End Sub

Private Sub cmdRecebimentos_Click()
If dbContas.Recordset.RecordCount = 0 Then Exit Sub
If dbContas.Recordset.EOF = True Then Exit Sub

Load frmConciliaRecebimentos
With frmConciliaRecebimentos
  With .dbPendencias
    .RecordSource = "select *from PrevisaoRecebimentos where codigoconta=" & dbContas.Recordset!codigoconta & " and confirmado=0 order by dataPrevista"
    .Refresh
  End With
End With
frmConciliaRecebimentos.Show vbModal


End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If DataGrid1.Col <> 0 Then
  DataGrid1.Col = 0
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case vbKeyReturn
    SendKeys Chr(vbKeyTab)
    KeyAscii = 0
End Select

End Sub

Private Sub Form_Load()
With dbContas
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from contas order by Descri"
  .Refresh
End With
With dbDespesaTipo
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from contasDespesas order by descri"
  .Refresh
End With
With dbConcilia
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .RecordSource = "select *from concilia where codigoconta=0"
  .Refresh
End With
With dbConciliaPendente
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbPendencias
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbTemp
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbDespesasLanc
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
With dbMovimentacao
  .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Caminho & ";Mode=ReadWrite;Persist Security Info=False"
  .Refresh
End With
txtDataLanc.Value = Date
txtData.Value = DateAdd("m", -1, Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtData_LostFocus()
If IsDate(txtData.Value) Then
  If cboContas.Text = "" Then Exit Sub
  With dbConcilia
    .RecordSource = "select *from concilia where data>=#" & DataInglesa(Str(txtData.Value)) & "# and codigoconta=" & dbContas.Recordset!codigoconta & " order by data, codigoconciliaconta"
    .Refresh
    .Refresh
    If .Recordset.EOF = False Then .Recordset.MoveLast
  End With
Else
  dbConcilia.RecordSource = "select *from concilia where codigoconta=" & dbContas.Recordset!codigoconta & " order by data, codigoconciliaconta"
  dbConcilia.Refresh
  dbConcilia.Recordset.MoveLast
End If
End Sub

Private Sub txtDataLanc_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataLanc_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataLanc_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub txtValor_GotFocus()
With txtValor
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case Asc(".")
    KeyAscii = 0
    SendKeys ","
End Select
End Sub

Private Sub txtValor_LostFocus()
If IsNumeric(txtValor.Text) = False Then
  txtValor.Text = "1"
End If
If dbDespesaTipo.Recordset.EOF = True Then Exit Sub
lblTotal.Caption = Format((CDbl(txtValor.Text) * dbDespesaTipo.Recordset!Valor), "currency")
End Sub
