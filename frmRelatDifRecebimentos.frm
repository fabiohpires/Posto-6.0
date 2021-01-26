VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatDifRecebimentos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Diferença de Recebimentos de Cartão"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
   Icon            =   "frmRelatDifRecebimentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc qPrevisaoRecebimentos 
      Height          =   330
      Left            =   2640
      Top             =   3360
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   $"frmRelatDifRecebimentos.frx":0442
      Caption         =   "qPrevisaoRecebimentos"
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
      Left            =   9600
      Picture         =   "frmRelatDifRecebimentos.frx":04C9
      Style           =   1  'Graphical
      TabIndex        =   15
      Tag             =   "Imprimir"
      Top             =   600
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbContas 
      Height          =   330
      Left            =   2640
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSAdodcLib.Adodc dbTipoRecebimento 
      Height          =   330
      Left            =   2640
      Top             =   2640
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select *from FormaDePagamento order by descri"
      Caption         =   "dbTipoRecebimento"
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
      Left            =   7440
      TabIndex        =   13
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   5760
      TabIndex        =   12
      Top             =   960
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo cboRecebimento 
      Bindings        =   "frmRelatDifRecebimentos.frx":0F4B
      Height          =   315
      Left            =   5400
      TabIndex        =   9
      Top             =   480
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Frame Frame1 
      Caption         =   " Data de Referência "
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.OptionButton Option3 
         Caption         =   "Data Recebida"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Data Prevista"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Data de Entrada"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc dbPrevisaoRecebimentos 
      Height          =   330
      Left            =   2640
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "select *from Cartoes"
      Caption         =   "dbPrevisaoRecebimentos"
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
      Bindings        =   "frmRelatDifRecebimentos.frx":0F6B
      Height          =   4695
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   8281
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "DataLanc"
         Caption         =   "Dt.Entrada"
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
         DataField       =   "DataPrevista"
         Caption         =   "Dt.Prevista"
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
         DataField       =   "DataRecebida"
         Caption         =   "Dt.Rec."
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
      BeginProperty Column03 
         DataField       =   "Conta"
         Caption         =   "Conta"
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
      BeginProperty Column05 
         DataField       =   "ValorBruto"
         Caption         =   "Valor Bruto"
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
         DataField       =   "ValorLiquido"
         Caption         =   "Liq.Previsto"
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
      BeginProperty Column07 
         DataField       =   "ValorRecebido"
         Caption         =   "V.Confirmado"
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
      BeginProperty Column08 
         DataField       =   "Diferenca"
         Caption         =   "Dif.Rec."
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
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   945,071
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1440
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   750,047
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   3960
      TabIndex        =   7
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   123535361
      CurrentDate     =   37767
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   2280
      TabIndex        =   5
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      Format          =   123535361
      CurrentDate     =   37767
   End
   Begin MSDataListLib.DataCombo cboConta 
      Bindings        =   "frmRelatDifRecebimentos.frx":0F90
      Height          =   315
      Left            =   2280
      TabIndex        =   11
      Top             =   1080
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "confirmado"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qPrevisaoRecebimentos"
      Height          =   255
      Left            =   7440
      TabIndex        =   23
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "dif"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qPrevisaoRecebimentos"
      Height          =   255
      Left            =   9000
      TabIndex        =   19
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "liquido"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qPrevisaoRecebimentos"
      Height          =   255
      Left            =   5880
      TabIndex        =   18
      Top             =   6600
      Width           =   1575
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "bruto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qPrevisaoRecebimentos"
      Height          =   255
      Left            =   4200
      TabIndex        =   17
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Diferença:"
      Height          =   255
      Left            =   9000
      TabIndex        =   22
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Previsto:"
      Height          =   255
      Left            =   5880
      TabIndex        =   21
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Bruto:"
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Total:"
      Height          =   255
      Left            =   3240
      TabIndex        =   16
      Top             =   6360
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Conta:"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo de Recebimento:"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   3720
      TabIndex        =   6
      Top             =   480
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "Período:"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Confirmado:"
      Height          =   255
      Left            =   7440
      TabIndex        =   24
      Top             =   6360
      Width           =   975
   End
End
Attribute VB_Name = "frmRelatDifRecebimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

StrTemp = "Relatório de Diferença de Recebimentos"
Printer.FontSize = 14
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp


Printer.FontSize = 8
StrTemp = "Data: " & Format(Dia, "long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboRecebimento.Text <> "" Then
  StrTemp = "Tipo de Recebimento: " & cboRecebimento.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

If cboConta.Text <> "" Then
  StrTemp = "Conta: " & cboConta.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Printer.CurrentY = Printer.CurrentY + 2

StrTemp = "Dt.Entrada"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Dt.Prevista"
Printer.CurrentX = 15
Printer.Print StrTemp;

StrTemp = "Dt.Recebida"
Printer.CurrentX = 30
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 45
Printer.Print StrTemp;

StrTemp = "Recebimento"
Printer.CurrentX = 75
Printer.Print StrTemp;

StrTemp = "V.Bruto"
Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Liq.Prev."
Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "V.Confirmado"
Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Dif. Rec."
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub CboConta_LostFocus()
If cboConta.Text = "" Then Exit Sub
With dbContas
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboConta.Text & "'"
  If .Recordset.EOF = False Then
    cboConta.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cboRecebimento_LostFocus()
If cboRecebimento.Text = "" Then Exit Sub
With dbTipoRecebimento
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "descri='" & cboRecebimento.Text & "'"
  If .Recordset.EOF = False Then
    cboRecebimento.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String, Ordem As String
Dim StrTemp2 As String

If Option1.Value = True Then
  StrTemp = "select *from Cartoes where confirmado=-1 and datalanc between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  StrTemp2 = "select sum(valorbruto) as bruto, sum(valorliquido) as liquido, sum(valorrecebido) as confirmado, sum(diferenca) as dif from Cartoes where confirmado=-1 and datalanc between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  Ordem = "datalanc"
End If
If Option2.Value = True Then
  StrTemp = "select *from Cartoes where confirmado=-1 and dataPrevista between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  StrTemp2 = "select sum(valorbruto) as bruto, sum(valorliquido) as liquido, sum(valorrecebido) as confirmado, sum(diferenca) as dif from Cartoes where confirmado=-1 and dataPrevista between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  Ordem = "dataprevista"
End If
If Option3.Value = True Then
  StrTemp = "select *from Cartoes where confirmado=-1 and datarecebida between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  StrTemp2 = "select sum(valorbruto) as bruto, sum(valorliquido) as liquido, sum(valorrecebido) as confirmado, sum(diferenca) as dif from Cartoes where confirmado=-1 and datarecebida between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  Ordem = "datarecebida"
End If

If cboRecebimento.Text <> "" Then
  If dbTipoRecebimento.Recordset.EOF = False Then
    If dbTipoRecebimento.Recordset!Descri = cboRecebimento.Text Then
      StrTemp = StrTemp & " and codigoformapg=" & dbTipoRecebimento.Recordset!CodigoPagamento
      StrTemp2 = StrTemp2 & " and codigoformapg=" & dbTipoRecebimento.Recordset!CodigoPagamento
    End If
  End If
End If

If cboConta.Text <> "" Then
  If dbContas.Recordset.EOF = False Then
    If dbContas.Recordset!Descri = cboConta.Text Then
      StrTemp = StrTemp & " and codigoconta=" & dbContas.Recordset!CodigoConta
      StrTemp2 = StrTemp2 & " and codigoconta=" & dbContas.Recordset!CodigoConta
    End If
  End If
End If

With dbPrevisaoRecebimentos
  .RecordSource = StrTemp
  .Refresh
  .Recordset.Sort = Ordem
End With
With qPrevisaoRecebimentos
  .RecordSource = StrTemp2
  .Refresh
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Dia As Date, Largura As Double
Dim Bruto As Currency, BrutoParcial As Currency
Dim Previsto As Currency, PrevistoParcial As Currency
Dim Confirmado As Currency, ConfirmadoParcial As Currency
Dim Diferenca As Currency, DiferencaParcial As Currency

With dbPrevisaoRecebimentos
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  .Recordset.MoveFirst
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Date
  
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = Format(BrutoParcial, "Currency")
      Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(PrevistoParcial, "Currency")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(ConfirmadoParcial, "Currency")
      Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(DiferencaParcial, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      BrutoParcial = 0
      PrevistoParcial = 0
      ConfirmadoParcial = 0
      DiferencaParcial = 0
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
      
    End If
    
    StrTemp = Format(.Recordset!DataLanc, "short date")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!DataPrevista, "Short date")
    Printer.CurrentX = 15
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!DataRecebida, "short date")
    Printer.CurrentX = 30
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Conta
    Printer.CurrentX = 45
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 75
    Printer.Print StrTemp;
    
    Bruto = Bruto + .Recordset!ValorBruto
    BrutoParcial = BrutoParcial + .Recordset!ValorBruto
    StrTemp = Format(.Recordset!ValorBruto, "Currency")
    Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Previsto = Previsto + .Recordset!valorliquido
    PrevistoParcial = PrevistoParcial + .Recordset!valorliquido
    StrTemp = Format(.Recordset!valorliquido, "Currency")
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Confirmado = Confirmado + .Recordset!ValorRecebido
    ConfirmadoParcial = ConfirmadoParcial + .Recordset!ValorRecebido
    StrTemp = Format(.Recordset!ValorRecebido, "Currency")
    Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Diferenca = Diferenca + .Recordset!Diferenca
    DiferencaParcial = DiferencaParcial + .Recordset!Diferenca
    StrTemp = Format(.Recordset!Diferenca, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(BrutoParcial, "Currency")
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(PrevistoParcial, "Currency")
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(ConfirmadoParcial, "Currency")
  Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(DiferencaParcial, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(Bruto, "Currency")
  Printer.CurrentX = 130 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(Previsto, "Currency")
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(Confirmado, "Currency")
  Printer.CurrentX = 170 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(Diferenca, "Currency")
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
If dbPrevisaoRecebimentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  dbPrevisaoRecebimentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  dbPrevisaoRecebimentos.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
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
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date

With dbPrevisaoRecebimentos
  .ConnectionString = CaminhoADO
  .Refresh
  .Recordset.Filter = "codigocartao=0"
End With
With qPrevisaoRecebimentos
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(valorbruto) as bruto, sum(valorliquidoprevisto) as liquido, sum(valorconfirmado) as confirmado, sum(difrecebido) as diferenca from previsaorecebimentos where codigoprevisaorecebe=0"
  .Refresh
End With
With dbContas
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbTipoRecebimento
  .ConnectionString = CaminhoADO
  .Refresh
End With

End Sub
