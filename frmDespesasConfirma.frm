VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "Msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDespesasConfirma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmar Despesas Lançadas"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11595
   Icon            =   "frmDespesasConfirma.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExibePagamentos 
      Caption         =   "Exibe Pagamentos"
      Height          =   495
      Left            =   2520
      TabIndex        =   28
      Top             =   5760
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo cboDespesaTipo2 
      Bindings        =   "frmDespesasConfirma.frx":0442
      Height          =   315
      Left            =   4560
      TabIndex        =   14
      Top             =   960
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc dbDespesaTipo2 
      Height          =   330
      Left            =   2400
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from despesatipo order by descri"
      Caption         =   "DbDespesaTipo"
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
   Begin VB.ComboBox cboPagoComo 
      Height          =   315
      ItemData        =   "frmDespesasConfirma.frx":045F
      Left            =   120
      List            =   "frmDespesasConfirma.frx":046C
      TabIndex        =   11
      Text            =   "Todas"
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cboContabilidade 
      Height          =   315
      ItemData        =   "frmDespesasConfirma.frx":048F
      Left            =   8040
      List            =   "frmDespesasConfirma.frx":049C
      TabIndex        =   9
      Text            =   "Todas"
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboConfirmado 
      Height          =   315
      ItemData        =   "frmDespesasConfirma.frx":04BF
      Left            =   5640
      List            =   "frmDespesasConfirma.frx":04CC
      TabIndex        =   7
      Text            =   "Todas"
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboAutoriazado 
      Height          =   315
      ItemData        =   "frmDespesasConfirma.frx":0507
      Left            =   3240
      List            =   "frmDespesasConfirma.frx":0517
      TabIndex        =   5
      Text            =   "Não Autorizadas"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   7920
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox txtDescri 
      Height          =   285
      Left            =   4680
      TabIndex        =   20
      Top             =   5280
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc DbDespesaTipo 
      Height          =   330
      Left            =   2400
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from despesatipo order by descri"
      Caption         =   "DbDespesaTipo"
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
   Begin MSDataListLib.DataCombo cboDespesatipo 
      Bindings        =   "frmDespesasConfirma.frx":0554
      Height          =   315
      Left            =   840
      TabIndex        =   18
      Top             =   5280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   9600
      TabIndex        =   15
      Top             =   840
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   161349633
      CurrentDate     =   38286
   End
   Begin VB.CheckBox chkFechadas 
      Caption         =   "Despesas já fechadas de meses anteriores"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   840
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc dbSoma 
      Height          =   330
      Left            =   2400
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select sum(valor) as total from despesaslanc2 where autorizacao=0 and fechamento=0"
      Caption         =   "dbSoma"
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
      Height          =   495
      Left            =   3840
      Picture         =   "frmDespesasConfirma.frx":0570
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "Imprimir"
      Top             =   5760
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbDespesaLanc 
      Height          =   330
      Left            =   2400
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Posto.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from DespesasLanc2 where codigofechamento=0 and autorizacao=0 order by data"
      Caption         =   "dbDespesaLanc"
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
      Height          =   495
      Left            =   10440
      TabIndex        =   27
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "&Confirmar"
      Height          =   495
      Left            =   120
      TabIndex        =   24
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelaConfirma 
      Caption         =   "Cancela Confirmação"
      Height          =   495
      Left            =   1200
      TabIndex        =   25
      Top             =   5760
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDespesasConfirma.frx":0FF2
      Height          =   3855
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Vencimento"
         Caption         =   "Venc."
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
         DataField       =   "Data"
         Caption         =   "Lanc."
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Obs"
         Caption         =   "Obs"
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
      BeginProperty Column04 
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
      BeginProperty Column05 
         DataField       =   "ValorPago"
         Caption         =   "Pago"
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
      BeginProperty Column06 
         DataField       =   "Autorizacao"
         Caption         =   "Autorizada"
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
      BeginProperty Column07 
         DataField       =   "Origem"
         Caption         =   "Origem"
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
      BeginProperty Column08 
         DataField       =   "DataContabilidade"
         Caption         =   "Para Contab. em"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2670,236
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2355,024
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   959,811
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   161349633
      CurrentDate     =   38286
   End
   Begin VB.Label Label11 
      Caption         =   "Despesa:"
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Pago Como:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "Envio para Contabilidade:"
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label8 
      Caption         =   "Confirmada Descrição:"
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Autorização:"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Descrição:"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   5280
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Despesa:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "a"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Período de lançamento:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
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
      DataSource      =   "dbSoma"
      Height          =   255
      Left            =   9600
      TabIndex        =   23
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   9120
      TabIndex        =   22
      Top             =   5280
      Width           =   405
   End
End
Attribute VB_Name = "frmDespesasConfirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Despesas"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp


Printer.FontSize = 8
StrTemp = "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "Short date")
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.FontSize = 8
StrTemp = "Data: " & Format(Dia, "Long date")
Printer.CurrentX = 0
Printer.Print StrTemp;


StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 2


StrTemp = "Venc."
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Lanc."
Printer.CurrentX = 15
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 30
Printer.Print StrTemp;

StrTemp = "Observações"
Printer.CurrentX = 80
Printer.Print StrTemp;

StrTemp = "Pago Como"
Printer.CurrentX = 140
Printer.Print StrTemp;

StrTemp = "Valor"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1


End Sub

Private Sub Filtrar()
Dim StrTemp As String, StrTemp2 As String
Dim StrDescri As String, strFechamento As String
Dim Contabilidade As String, strPagoComo As String

If chkFechadas.Value = vbChecked Then
  strFechamento = ""
Else
  strFechamento = " and fechamento=0"
End If

Select Case cboConfirmado
  Case "Não Confirmada Descrição"
    StrDescri = " and usuario='Não'"
  Case "Confirmada Descrição"
    StrDescri = " and usuario='Sim'"
  Case "Todas"
    StrDescri = ""
End Select

Select Case cboContabilidade.Text
  Case "Não Enviadas"
    Contabilidade = " and paracontabilidade=0"
  Case "Enviadas"
    Contabilidade = " and paracontabilidade=-1"
  Case "Todas"
    Contabilidade = ""
End Select

Select Case cboPagoComo.Text
  Case "Todas"
    strPagoComo = ""
  Case Else
    strPagoComo = " and pagarcomo='" & cboPagoComo.Text & "'"
End Select

StrTemp = "select *from despesaslanc2 where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and fechamentodiario=-1" & strFechamento & StrDescri & Contabilidade & strPagoComo
StrTemp2 = "select sum(valor) as total from despesaslanc2 where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and fechamentodiario=-1" & strFechamento & StrDescri & Contabilidade & strPagoComo

If cboDespesaTipo2.Text <> "" Then
  If dbDespesaTipo2.Recordset.EOF = False And dbDespesaTipo2.Recordset.BOF = False Then
    StrTemp = StrTemp & " and codigodespesa=" & dbDespesaTipo2.Recordset!CodigoDespesa
    StrTemp2 = StrTemp2 & " and codigodespesa=" & dbDespesaTipo2.Recordset!CodigoDespesa
  End If
End If

Select Case cboAutoriazado.Text
  Case "Não Autorizadas"
    StrTemp = StrTemp & " and autorizacao=0 and produto=0 order by data, hora"
    StrTemp2 = StrTemp2 & " and autorizacao=0 and produto=0"
  Case "Autorizadas"
    StrTemp = StrTemp & " and autorizacao=-1 and produto=0 order by data, hora"
    StrTemp2 = StrTemp2 & " and autorizacao=-1 and produto=0"
  Case "Todas"
    StrTemp = StrTemp & " and produto=0 order by data, hora"
    StrTemp2 = StrTemp2 & " and produto=0"
  Case "Despesas Bancárias"
    StrTemp = StrTemp & " and produto=0 and origem='Conciliação' order by data, hora"
    StrTemp2 = StrTemp2 & " and produto=0 and origem='Conciliação'"
End Select

With dbDespesaLanc
  .ConnectionString = CaminhoADO
  .RecordSource = StrTemp
  .Refresh
End With
With dbSoma
  .ConnectionString = CaminhoADO
  .RecordSource = StrTemp2
  .Refresh
End With


End Sub

Private Sub cboAutoriazado_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboAutoriazado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboAutoriazado_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cboConfirmado_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboConfirmado_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboConfirmado_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cboContabilidade_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cboContabilidade_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub cboContabilidade_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cboDespesatipo_LostFocus()
With dbDespesaTipo
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  If cboDespesatipo.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboDespesatipo.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  cboDespesatipo.Text = .Recordset!Descri
End With
End Sub

Private Sub cboDespesaTipo2_LostFocus()
With dbDespesaTipo2
  .Refresh
  If .Recordset.EOF = True Then Exit Sub
  If cboDespesaTipo2.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboDespesaTipo2.Text & "'"
  If .Recordset.EOF = True Then Exit Sub
  cboDespesaTipo2.Text = .Recordset!Descri
End With
End Sub

Private Sub cmdAlterar_Click()
If dbDespesaLanc.Recordset.EOF = True Then
  MsgBox "Selecione uma despesa primeiro!"
  Exit Sub
End If
If dbDespesaLanc.Recordset!Fechamento = True Then
  MsgBox "Despesa que já feito o fechamento não pode ser alterada!"
  Exit Sub
End If
If dbDespesaLanc.Recordset!autorizacao = True Then
  MsgBox "Despesa já confirmada pela administração!"
  Exit Sub
End If
If cboDespesatipo.Text = "" Then
  MsgBox "Informe um tipo de despesa!"
  cboDespesatipo.SetFocus
  Exit Sub
End If
If dbDespesaTipo.Recordset.EOF = True Then
  MsgBox "Cadastro de tipo de despesas vazio!"
  Exit Sub
End If
If dbDespesaTipo.Recordset!Descri <> cboDespesatipo.Text Then
  MsgBox "Despesa inválida!"
  cboDespesatipo.SetFocus
  Exit Sub
End If
If txtDescri.Text = "" Then
  MsgBox "Informe uma descrição para a despesa!"
  txtDescri.SetFocus
  Exit Sub
End If

With dbDespesaLanc
  .Recordset!CodigoDespesa = dbDespesaTipo.Recordset!CodigoDespesa
  .Recordset!Descri = dbDespesaTipo.Recordset!Descri
  .Recordset!Obs = txtDescri.Text
  .Recordset!Usuario = "Sim"
  
  .Recordset.Update
End With

End Sub

Private Sub cmdCancelaConfirma_Click()
Dim CodigoDespesa As Double
With dbDespesaLanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  CodigoDespesa = .Recordset.AbsolutePosition
  .Recordset!autorizacao = False
  .Recordset.Update
  .Refresh
  .Refresh
  On Error Resume Next
  .Recordset.AbsolutePage = CodigoDespesa
End With
End Sub

Private Sub cmdConfirmar_Click()
Dim CodigoDespesa As Double
With dbDespesaLanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  If .Recordset.EOF = True Then Exit Sub
  CodigoDespesa = .Recordset.AbsolutePosition
  .Recordset!autorizacao = True
  .Recordset.Update
  .Refresh
  .Refresh
  .Refresh
  On Error Resume Next
  .Recordset.AbsolutePosition = CodigoDespesa
End With
End Sub

Private Sub cmdExibePagamentos_Click()
frmDespesasConfirmaExibePg.Show
frmDespesasConfirmaExibePg.Exibir dbDespesaLanc.Recordset!CodigoDespesaLanc
frmDespesasConfirmaExibePg.Hide
frmDespesasConfirmaExibePg.Show vbModal
End Sub

Private Sub cmdExibir_Click()
Filtrar
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Total As Currency, SubTotal As Currency, TipoQuebra As String, Quebra As String

With dbDespesaLanc
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  TipoQuebra = .Recordset.Sort
  If TipoQuebra = "" Then
    TipoQuebra = "Data"
  End If
  Quebra = .Recordset(TipoQuebra)
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Largura = 190
  Dia = Now
  
  
  Printer.FontName = "Arial"
  Printer.ScaleMode = vbMillimeters
  Printer.DrawWidth = 2
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Total: " & Format(Total, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    If Quebra <> .Recordset(TipoQuebra) Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Sub-Total: " & Format(SubTotal, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      SubTotal = 0
      
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      Quebra = .Recordset(TipoQuebra)
    End If
    Printer.FontSize = 8
    If IsNull(.Recordset!Vencimento) = False Then
      StrTemp = Format(.Recordset!Vencimento, "short date")
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Data, "short date")
    Printer.CurrentX = 15
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 30
    Printer.Print StrTemp;
    
    If IsNull(.Recordset!Obs) = False Then
      StrTemp = .Recordset!Obs
      Printer.CurrentX = 80
      Printer.Print StrTemp;
    End If
    If IsNull(.Recordset!pagarcomo) = False Then
      StrTemp = .Recordset!pagarcomo
    Else
      StrTemp = ""
    End If
    Printer.CurrentX = 140
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Valor, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    Total = Total + .Recordset!Valor
    SubTotal = SubTotal + .Recordset!Valor
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = "Sub-Total: " & Format(SubTotal, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  StrTemp = "Total: " & Format(Total, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.EndDoc
End With
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
dbDespesaLanc.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
Select Case KeyAscii
  Case vbKeyReturn
    KeyAscii = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub Form_Load()
Dim db As New ADODB.Connection
Dim dbTemp As New ADODB.Recordset

db.Open CaminhoADO
dbTemp.CursorLocation = adUseClient
dbTemp.Open "Select despesaslanc2.PagarComo from despesaslanc2 group by PagarComo order by PagarComo", db, adOpenKeyset, adLockOptimistic
With cboPagoComo
  .Clear
  .AddItem "Todas"
  If dbTemp.RecordCount <> 0 Then
    Do While dbTemp.EOF = False
      If IsNull(dbTemp!pagarcomo) = False Then .AddItem dbTemp!pagarcomo
      dbTemp.MoveNext
    Loop
  End If
  .Text = "Todas"
End With
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbDespesaTipo
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from despesaslanc2 where usuario='Não'"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    .RecordSource = "Select *from despesaslanc2"
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      .Recordset.MoveLast
      .Recordset.MoveFirst
      Do While .Recordset.EOF = False
        .Recordset!Usuario = "Não"
        .Recordset.Update
        .Recordset.MoveNext
      Loop
    End If
  End If
  .RecordSource = "select *from despesatipo order by descri"
  .Refresh
End With
With dbDespesaTipo2
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from despesatipo order by descri"
  .Refresh
End With



Filtrar
Select Case Usuarios.Grupo.AdmConfirma
  Case 1 'Somente leitura
    cmdConfirmar.Enabled = False
    cmdCancelaConfirma.Enabled = False
  Case 2 'Liberado
    
End Select

End Sub

Private Sub optAutoriza_Click(Index As Integer)
Filtrar
End Sub

Private Sub optDescriNao_Click()
Filtrar
End Sub

Private Sub optDescriSim_Click()
Filtrar
End Sub

Private Sub optDescriTodas_Click()
Filtrar
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

Private Sub txtDataIni_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
