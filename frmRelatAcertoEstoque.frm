VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatAcertoEstoque 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Acerto de Estoque"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10725
   Icon            =   "frmRelatAcertoEstoque.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6600
      Picture         =   "frmRelatAcertoEstoque.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "Imprimir"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   975
   End
   Begin VB.OptionButton optCombustivel 
      Caption         =   "Combustível"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton optNaoCombustivel 
      Caption         =   "Não Combustível"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.OptionButton optAmbos 
      Caption         =   "Ambos"
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   840
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc qEntrada 
      Height          =   330
      Left            =   1920
      Top             =   4200
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
      RecordSource    =   "select sum(ValorUtilizado) as qtd, sum(Valordiferenca) as total from produtosacerto"
      Caption         =   "qEntrada"
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
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   1920
      Top             =   3840
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
      RecordSource    =   "select *from produtos order by descri"
      Caption         =   "dbProdutos"
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
   Begin MSDataListLib.DataCombo cboProduto 
      Bindings        =   "frmRelatAcertoEstoque.frx":0EC4
      Height          =   315
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin MSAdodcLib.Adodc dbEntrada 
      Height          =   330
      Left            =   1920
      Top             =   3480
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
      RecordSource    =   "select *from produtosacerto order by datalancada"
      Caption         =   "dbEntrada"
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
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   72351745
      CurrentDate     =   37678
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatAcertoEstoque.frx":0EDD
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6800
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
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "DataLancada"
         Caption         =   "DataLancada"
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
         DataField       =   "CodigoProduto"
         Caption         =   "Cod."
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
         DataField       =   "Descri"
         Caption         =   "Descri"
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
         DataField       =   "EstoqueAnterior"
         Caption         =   "Est. Antes"
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
      BeginProperty Column04 
         DataField       =   "EstoquePosterior"
         Caption         =   "Est. Depois"
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
      BeginProperty Column05 
         DataField       =   "EstoqueResultante"
         Caption         =   "Est. Final"
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
      BeginProperty Column06 
         DataField       =   "ValorUtilizado"
         Caption         =   "Qtd"
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
         DataField       =   "PrecoCompra"
         Caption         =   "Preço Compra"
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
         DataField       =   "PrecoVenda"
         Caption         =   "Preço Venda"
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
      BeginProperty Column09 
         DataField       =   "ValorDiferenca"
         Caption         =   "Diferença"
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
            ColumnWidth     =   1214,929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   689,953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   884,976
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   854,929
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   555,024
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnWidth     =   1035,213
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   854,929
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   16
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   15
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      Height          =   195
      Left            =   7680
      TabIndex        =   13
      Top             =   5280
      Width           =   405
   End
   Begin VB.Label Label6 
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
      DataSource      =   "qEntrada"
      Height          =   255
      Left            =   8280
      TabIndex        =   12
      Top             =   5280
      Width           =   1125
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
      DataSource      =   "qEntrada"
      Height          =   255
      Left            =   9360
      TabIndex        =   11
      Top             =   5280
      Width           =   1245
   End
End
Attribute VB_Name = "frmRelatAcertoEstoque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOrdem As String

Private Sub Cabeca(ByVal Largura As Double, ByVal Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Acerto de Estoque"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

StrTemp = NomePosto
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 10

StrTemp = "Data: " & Format(Dia, "Long Date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " '& Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Período: " & Format(txtDataIni.Value, "Short Date") & " a " & Format(txtDataFim.Value, "Short Date")
Printer.CurrentX = 0
Printer.Print StrTemp

If cboProduto.Text = dbProdutos.Recordset!Descri Then
  StrTemp = "Código: " & txtCodigo.Text & " - Produto: " & cboProduto.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If
Printer.CurrentY = Printer.CurrentY + 1


StrTemp = "Data"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Código"
Printer.CurrentX = 43 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Descrição"
Printer.CurrentX = 45
Printer.Print StrTemp;

StrTemp = "Quantidade"
Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Diferença"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub Filtrar()
Dim StrTemp As String, StrTemp2 As String
StrTemp = "Select *from produtosacerto where datalancada between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
StrTemp2 = "select sum(ValorUtilizado) as qtd, sum(Valordiferenca) as total from produtosacerto where datalancada between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
If cboProduto.Text = dbProdutos.Recordset!Descri Then
  StrTemp = StrTemp & " and codproduto=" & dbProdutos.Recordset!CodigoProduto
  StrTemp2 = StrTemp2 & " and codproduto=" & dbProdutos.Recordset!CodigoProduto
Else
  If optAmbos.Value = False Then
    If optCombustivel.Value = True Then
      StrTemp = StrTemp & " and combustivel=-1"
      StrTemp2 = StrTemp2 & " and combustivel=-1"
    Else
      StrTemp = StrTemp & " and combustivel=0"
      StrTemp2 = StrTemp2 & " and combustivel=0"
    End If
  End If
End If

With dbEntrada
  .RecordSource = StrTemp
  .Refresh
  .Refresh
  .Recordset.Sort = strOrdem
End With
qEntrada.RecordSource = StrTemp2
qEntrada.Refresh
qEntrada.Refresh

End Sub
Private Sub cboProduto_LostFocus()
Me.KeyPreview = True
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If cboProduto.Text = "" Then Exit Sub
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

Private Sub cmdExibir_Click()
Filtrar
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date
Dim Total1 As Double, Total2 As Currency

With dbEntrada
  If .Recordset.RecordCount = 0 Then Exit Sub
  
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 155
  Dia = Now
  
  Cabeca Largura, Dia
  
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight + 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = "Sub-Total: " & Total1
      Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(Total2, "Currency")
      Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    
    StrTemp = Format(.Recordset!datalancada, "Short date")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodigoProduto
    Printer.CurrentX = 43 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 45
    Printer.Print StrTemp;
    
    Total1 = Total1 + .Recordset!Valorutilizado
    StrTemp = .Recordset!Valorutilizado
    Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Total2 = Total2 + .Recordset!ValorDiferenca
    StrTemp = Format(.Recordset!ValorDiferenca, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = "Total: " & Total1
  Printer.CurrentX = 125 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(Total2, "Currency")
  Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
End With
Printer.EndDoc
NaoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If strOrdem = DataGrid1.Columns(ColIndex).DataField Then
  strOrdem = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  strOrdem = DataGrid1.Columns(ColIndex).DataField
End If
Filtrar
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
strOrdem = "datalancada"
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbEntrada
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .Refresh
End With
With qEntrada
  .ConnectionString = CaminhoADO
  .Refresh
End With
Filtrar
End Sub


Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With dbProdutos
  .Refresh
  If .Recordset.RecordCount = 0 Then Exit Sub
  If IsNumeric(txtCodigo.Text) = False Then Exit Sub
  .Recordset.Find "codigo=" & txtCodigo.Text
  If .Recordset.EOF = False Then
    txtCodigo.Text = .Recordset!Codigo
    cboProduto.Text = .Recordset!Descri
  End If
End With
End Sub

