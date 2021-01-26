VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatVendasMedia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Total de Vendas"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmRelatVendasMedia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc qVendas 
      Height          =   330
      Left            =   1920
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
      CursorType      =   2
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
      RecordSource    =   $"frmRelatVendasMedia.frx":0442
      Caption         =   "qVendas"
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
   Begin MSAdodcLib.Adodc qVendasTotal 
      Height          =   330
      Left            =   1920
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      RecordSource    =   "SELECT Sum(qvendas.valortotal) AS total, Sum(qvendas.Quantidade) AS qtd From qVendas"
      Caption         =   "qVendasTotal"
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
      Left            =   120
      TabIndex        =   3
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   7080
      Picture         =   "frmRelatVendasMedia.frx":0535
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox txtCodFun 
      Height          =   285
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   5
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   37678
   End
   Begin MSAdodcLib.Adodc dbFuncionarios 
      Height          =   330
      Left            =   1920
      Top             =   4440
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
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
      RecordSource    =   "select *from vendedores order by nome"
      Caption         =   "dbFuncionarios"
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
   Begin MSDataListLib.DataCombo cboFuncionario 
      Bindings        =   "frmRelatVendasMedia.frx":0FB7
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   360
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatVendasMedia.frx":0FD4
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   9551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CodProduto"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
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
         DataField       =   "qtd"
         Caption         =   "Quantidade"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "total"
         Caption         =   "Total"
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
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3420,284
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   15
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Vendido:"
      Height          =   195
      Left            =   4920
      TabIndex        =   13
      Top             =   6360
      Width           =   1245
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   12
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Funcionário:"
      Height          =   195
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label7 
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
         SubFormatType   =   2
      EndProperty
      DataSource      =   "qVendasTotal"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "qtd"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      DataSource      =   "qVendasTotal"
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade Vendido:"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   6360
      Width           =   1725
   End
End
Attribute VB_Name = "frmRelatVendasMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Indice As String

Private Sub Cabeca(ByVal Dia As Date, ByVal Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Média de Vendas"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.CurrentY = 0
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8
StrTemp = "Data: " & Format(Dia, "long date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "Short date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboFuncionario.Text <> "" Then
  StrTemp = "Código do Funcionário: " & txtCodFun.Text & "    Funcionário: " & cboFuncionario.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Printer.Print ""

StrTemp = "Codigo"
Printer.CurrentX = 28 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 30
Printer.Print StrTemp;

StrTemp = "Vendido"
Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total Vendido"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub cboFuncionario_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = VerificaTecla(KeyCode)
End Sub

Private Sub cboFuncionario_LostFocus()
With dbFuncionarios
  .Refresh
  If cboFuncionario.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "nome='" & cboFuncionario.Text & "'"
  If .Recordset.EOF = False Then
    txtCodFun.Text = .Recordset!Codigo
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub


Private Sub cmdExibir_Click()
Dim StrTemp As String
Dim StrTemp2 As String

StrTemp = "SELECT qvendas.CodProduto, qvendas.Descri, Sum(venda2.valortotal) AS total, Sum(qvendas.Quantidade) AS qtd From qVendas where venda2.data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
StrTemp2 = "SELECT Sum(venda2.valortotal) AS total, Sum(qvendas.Quantidade) AS qtd From qVendas where venda2.data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"

If cboFuncionario.Text <> "" Then
  Call cboFuncionario_LostFocus
  StrTemp = StrTemp & " and codigo=" & dbFuncionarios.Recordset!Codigo
  StrTemp2 = StrTemp2 & " and codigo=" & dbFuncionarios.Recordset!Codigo
End If
StrTemp = StrTemp & " GROUP BY qvendas.CodProduto, qvendas.Descri order by codproduto"
qVendas.RecordSource = StrTemp
qVendas.Refresh
qVendasTotal.RecordSource = StrTemp2
qVendasTotal.Refresh

End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double
Dim Qtd As Double, QtdParcial As Double
Dim VTotal As Currency, VParcial As Currency
Dim Comissao As Currency, ComissaoParcial As Currency
Dim Quebra As String

With qVendas
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe registro a ser impresso!"
    Exit Sub
  End If
  
  On Error GoTo TrataErro
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  
  Cabeca Dia, Largura
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = Format(QtdParcial, "#,##0")
      Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(VParcial, "#,##0.00")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Dia, Largura
    End If
    
    StrTemp = .Recordset!CodProduto
    Printer.CurrentX = 28 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 30
    Printer.Print StrTemp;
    
    
    Qtd = Qtd + .Recordset!Qtd
    QtdParcial = QtdParcial + .Recordset!Qtd
    StrTemp = .Recordset!Qtd
    Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    VTotal = VTotal + .Recordset("total")
    VParcial = VParcial + .Recordset("total")
    StrTemp = Format(.Recordset("total"), "#,##0.00")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(QtdParcial, "#,##0")
  Printer.CurrentX = 150 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(VParcial, "#,##0.00")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
End With
Printer.EndDoc
TrataErro:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
Indice = DataGrid1.Columns(ColIndex).DataField
If qVendas.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  qVendas.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  qVendas.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = VerificaTecla(KeyAscii)
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With qVendasTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "SELECT qvendas.CodProduto, qvendas.Descri, Sum(venda2.valortotal) AS total, Sum(qvendas.Quantidade) AS qtd From qVendas GROUP BY qvendas.CodProduto, qvendas.Descri order by codproduto"
  .Refresh
End With
With qVendasTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "SELECT Sum(valortotal) AS total, Sum(Quantidade) AS qtd From qVendas where codigovenda=0"
  .Refresh
End With
With dbFuncionarios
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from vendedores order by nome"
  .Refresh
End With
Call cmdExibir_Click
End Sub

Private Sub txtCodFun_GotFocus()
With txtCodFun
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodFun_LostFocus()
With dbFuncionarios
  .Refresh
  If txtCodFun.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "codigo=" & txtCodFun.Text
  If .Recordset.EOF = False Then
    txtCodFun.Text = .Recordset!Codigo
    cboFuncionario.Text = .Recordset!Nome
  End If
End With
End Sub


Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    txtCodFun.SetFocus
End Select
End Sub

Private Sub txtDataIni_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    txtDataFim.SetFocus
End Select
End Sub

Private Function VerificaTecla(ByVal Tecla As Integer) As Integer
Select Case Tecla
  Case vbKeyReturn
    VerificaTecla = 0
    SendKeys Chr(vbKeyTab)
  Case Else
    VerificaTecla = Tecla
End Select
End Function

