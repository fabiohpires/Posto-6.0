VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatProtestos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Protestos"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   Icon            =   "frmRelatProtestos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodClientesNotas 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin MSDataListLib.DataCombo cboClientesNotas 
      Bindings        =   "frmRelatProtestos.frx":0442
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.OptionButton optCobranca 
      Caption         =   "Cobranças"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   1095
   End
   Begin VB.OptionButton optCheques 
      Caption         =   "Cheques"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.OptionButton optTodos 
      Caption         =   "Todos"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin MSAdodcLib.Adodc qProtestos 
      Height          =   330
      Left            =   1920
      Top             =   4200
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
      RecordSource    =   "Select *from Protestos order by codigoprotesto"
      Caption         =   "qProtestos"
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
   Begin MSAdodcLib.Adodc dbClientesNotas 
      Height          =   330
      Left            =   1920
      Top             =   4920
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
      RecordSource    =   "Select *from clientes order by nome"
      Caption         =   "dbClientesNotas"
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
   Begin MSAdodcLib.Adodc dbChequesClientes 
      Height          =   330
      Left            =   1920
      Top             =   4560
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
      RecordSource    =   "Select *from chequesclientes order by nome"
      Caption         =   "dbChequesClientes"
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
      Left            =   10920
      Picture         =   "frmRelatProtestos.frx":0460
      Style           =   1  'Graphical
      TabIndex        =   16
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbProtestos 
      Height          =   330
      Left            =   1920
      Top             =   3840
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
      RecordSource    =   "Select *from Protestos order by codigoprotesto"
      Caption         =   "dbProtestos"
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
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   360
      Width           =   1095
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
      Format          =   72351745
      CurrentDate     =   39647
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
      Format          =   72351745
      CurrentDate     =   39647
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatProtestos.frx":0EE2
      Height          =   5535
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9763
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "DataLanc"
         Caption         =   "Lançado"
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
      BeginProperty Column01 
         DataField       =   "DataDocumento"
         Caption         =   "Vencimento"
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
      BeginProperty Column02 
         DataField       =   "TipoDocumento"
         Caption         =   "Tipo"
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
         DataField       =   "Status"
         Caption         =   "Status"
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
         DataField       =   "Valor"
         Caption         =   "Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
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
            Alignment       =   1
            ColumnWidth     =   1244,976
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1349,858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1289,764
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1379,906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   4094,929
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1409,953
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCodChequesClientes 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4440
      TabIndex        =   8
      Top             =   360
      Width           =   615
   End
   Begin MSDataListLib.DataCombo cboChequeClientes 
      Bindings        =   "frmRelatProtestos.frx":0EFC
      Height          =   315
      Left            =   5160
      TabIndex        =   10
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin VB.Label lblClientesNotas 
      Caption         =   "Clientes de Nota:"
      Height          =   255
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblCodigoNotas 
      Caption         =   "Código:"
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   19
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   9360
      TabIndex        =   18
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Período Lançado:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label lblChequesClientes 
      Caption         =   "Clientes de Cheques:"
      Height          =   255
      Left            =   5160
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblCodigoCheques 
      Caption         =   "Código:"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatProtestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTipo As String, strOrdem As String

Private Sub AtivaChequesClientes(ByVal Cheques As Boolean, ByVal Boletos As Boolean)
txtCodChequesClientes.Visible = Cheques
cboChequeClientes.Visible = Cheques
lblChequesClientes.Visible = Cheques
lblCodigoCheques.Visible = Cheques
txtCodClientesNotas.Visible = Boletos
cboClientesNotas.Visible = Boletos
lblClientesNotas.Visible = Boletos
lblCodigoNotas.Visible = Boletos
End Sub

Private Sub cboChequeClientes_LostFocus()
With dbChequesClientes
  .Refresh
  If cboChequeClientes.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "nome='" & cboChequeClientes.Text & "'"
  If .Recordset.EOF = False Then
    txtCodChequesClientes.Text = .Recordset!codigochequecliente
  End If
End With
End Sub

Private Sub cboClientesNotas_LostFocus()
With dbClientesNotas
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboClientesNotas.Text = "" Then Exit Sub
  .Recordset.Find "nome='" & cboClientesNotas.Text & "'"
  If .Recordset.EOF = False Then
    txtCodClientesNotas.Text = .Recordset!CodigoCliente
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrTemp As String, StrTemp2 As String

StrTemp = "Select *from protestos where datalanc between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
StrTemp2 = "select sum(valor) as total from protestos where datalanc between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
Select Case strTipo
  Case "Cheques"
    StrTemp = StrTemp & " and tipodocumento='Cheque'"
    If cboChequeClientes.Text <> "" Then
      With dbChequesClientes
        If .Recordset.EOF = False And .Recordset.BOF = False Then
          If .Recordset!Nome = cboChequeClientes.Text Then
            StrTemp = StrTemp & " and codigoclientecheque=" & .Recordset!codigochequecliente
            StrTemp2 = StrTemp2 & " and codigoclientecheque=" & .Recordset!codigochequecliente
          End If
        End If
      End With
    End If
  Case "Cobranças"
    StrTemp = StrTemp & " and tipodocumento='Cobrança'"
    If cboClientesNotas.Text <> "" Then
      With dbClientesNotas
        If .Recordset.EOF = False And .Recordset.BOF = False Then
          If .Recordset!Nome = cboClientesNotas.Text Then
            StrTemp = StrTemp & " and codigoclientenota=" & .Recordset!CodigoCliente
            StrTemp2 = StrTemp2 & " and codigoclientenota=" & .Recordset!CodigoCliente
          End If
        End If
      End With
    End If
End Select

With dbProtestos
  .RecordSource = StrTemp & strOrdem
  .Refresh
End With
With qProtestos
  .RecordSource = StrTemp2
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblTotal.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotal.Caption = Format(0, "Currency")
  End If
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String

On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then Exit Sub
On Error GoTo 0

StrTemp = Chr(vbKeyReturn) & "Impresso em: " & Format(Date, "long Date")
StrTemp = StrTemp & Chr(vbKeyReturn) & "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")

Select Case strTipo
  Case "Cheques"
    StrTemp = StrTemp & Chr(vbKeyReturn) & " Tipodo: Cheque"
    If cboChequeClientes.Text <> "" Then
      With dbChequesClientes
        If .Recordset.EOF = False And .Recordset.BOF = False Then
          StrTemp = StrTemp & "  Codigo: " & .Recordset!codigochequecliente & "  Nome: " & .Recordset!Nome
        End If
      End With
    End If
  Case "Cobranças"
    StrTemp = StrTemp & Chr(vbKeyReturn) & " Tipodo: Cobrança"
    If cboClientesNotas.Text <> "" Then
      With dbClientesNotas
        If .Recordset.EOF = False And .Recordset.BOF = False Then
          If .Recordset!Nome = cboClientesNotas.Text Then
            StrTemp = StrTemp & "  Codigo: " & .Recordset!CodigoCliente & "  Nome: " & .Recordset!Nome
          End If
        End If
      End With
    End If
End Select

ImprimeADOGrid DataGrid1, Printer, dbProtestos, 5, , , , , , "Protestos", NomePosto, StrTemp

Printer.EndDoc

If dbProtestos.Recordset.RecordCount <> 0 Then
  dbProtestos.Recordset.MoveFirst
End If

NaoImprime:

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField Then
  strOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = " order by " & DataGrid1.Columns(ColIndex).DataField
End If
Call cmdExibir_Click
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
strTipo = "Todos"
strOrdem = " order by DataLanc"
With dbProtestos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from protestos where codigoprotesto=0"
  .Refresh
End With
With qProtestos
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from protestos where codigoprotesto=0"
  .Refresh
End With
With dbChequesClientes
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbClientesNotas
  .ConnectionString = CaminhoADO
  .Refresh
End With
Call cmdExibir_Click
End Sub

Private Sub optCheques_Click()
If optCheques.Value = True Then strTipo = "Cheques"
AtivaChequesClientes True, False
End Sub

Private Sub optCobranca_Click()
If optCobranca.Value = True Then strTipo = "Cobranças"
AtivaChequesClientes False, True
End Sub

Private Sub optTodos_Click()
If optTodos.Value = True Then strTipo = "Todos"
AtivaChequesClientes False, False
End Sub

Private Sub txtCodChequesClientes_LostFocus()
With dbChequesClientes
  .Refresh
  If IsNumeric(txtCodChequesClientes.Text) = False Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.Find "codigochequecliente=" & txtCodChequesClientes.Text
  If .Recordset.EOF = False Then
    cboChequeClientes.Text = .Recordset!Nome
  End If
End With
End Sub

Private Sub txtCodClientesNotas_LostFocus()
With dbClientesNotas
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodClientesNotas.Text) = False Then Exit Sub
  .Recordset.Find "codigocliente=" & txtCodClientesNotas.Text
  If .Recordset.EOF = False Then
    cboClientesNotas.Text = .Recordset!Nome
  End If
End With
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
Select Case KeyCode
  Case vbKeyReturn
    KeyCode = 0
    SendKeys Chr(vbKeyTab)
End Select
End Sub

Private Sub txtDataIni_LostFocus()
Me.KeyPreview = True
End Sub
