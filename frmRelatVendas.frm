VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatVendas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Vendas"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   Icon            =   "frmRelatVendas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   9480
      Picture         =   "frmRelatVendas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "Imprimir"
      Top             =   720
      Width           =   735
   End
   Begin VB.CheckBox chkNaoFinalizados 
      Caption         =   "Caixas não confirmados também"
      Height          =   255
      Left            =   6840
      TabIndex        =   10
      Top             =   600
      Width           =   2655
   End
   Begin VB.Data dbVendas2 
      Caption         =   "dbVendas2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from venda2"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data qVendas2 
      Caption         =   "qVendas2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from qcomissoes2 where venda2.codigofechamento=0"
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmRelatVendas.frx":0EC4
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frmRelatVendas.frx":0EDB
      TabIndex        =   20
      Top             =   3960
      Width           =   10095
   End
   Begin MSAdodcLib.Adodc dbDespesaLanc 
      Height          =   330
      Left            =   4920
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
      RecordSource    =   "select *from despesaslanc2"
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
   Begin VB.CommandButton cmdPagar 
      Caption         =   "Pagar Comissões"
      Height          =   375
      Left            =   5400
      TabIndex        =   16
      Top             =   960
      Width           =   1575
   End
   Begin VB.CheckBox chkPago 
      Caption         =   "Permitir exibição das comissões pagas"
      Height          =   255
      Left            =   6840
      TabIndex        =   8
      Top             =   120
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc qVendasTotal 
      Height          =   330
      Left            =   1920
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
      RecordSource    =   "select sum(quantidade) as qtd, sum(venda.valortotal) as Vendido, sum(valorcomissao) as comissao from qcomissoes"
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
   Begin VB.TextBox txtCodFun 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   3000
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin MSDataListLib.DataCombo cboProduto 
      Bindings        =   "frmRelatVendas.frx":262A
      Height          =   315
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Descri"
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   64290817
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   64290817
      CurrentDate     =   37678
   End
   Begin MSAdodcLib.Adodc dbFuncionarios 
      Height          =   330
      Left            =   4920
      Top             =   1920
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
   Begin MSAdodcLib.Adodc dbProdutos 
      Height          =   330
      Left            =   1920
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
   Begin MSAdodcLib.Adodc qVendas 
      Height          =   330
      Left            =   1920
      Top             =   1920
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
      RecordSource    =   "select *from venda2"
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
   Begin MSDataListLib.DataCombo cboFuncionario 
      Bindings        =   "frmRelatVendas.frx":2643
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Nome"
      Text            =   ""
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmRelatVendas.frx":2660
      Height          =   2055
      Left            =   120
      TabIndex        =   19
      Top             =   1440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3625
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
         DataField       =   "Data"
         Caption         =   "Dt. Venda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/mm/yy ddd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CodProduto"
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
         DataField       =   "Quantidade"
         Caption         =   "Qtd."
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
         DataField       =   "ValorUnitario"
         Caption         =   "V.Unitário"
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
         DataField       =   "ValorTotal"
         Caption         =   "Total"
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
            Alignment       =   1
            ColumnWidth     =   599,811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4694,74
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   480,189
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   824,882
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   824,882
         EndProperty
      EndProperty
   End
   Begin VB.CheckBox chkExibirTodas 
      Caption         =   "Exibir todas as vendas"
      Height          =   255
      Left            =   6840
      TabIndex        =   9
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Vendido:"
      Height          =   195
      Left            =   7440
      TabIndex        =   30
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label lblTotalVendido1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   8760
      TabIndex        =   29
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label lblQtd1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade Vendido:"
      Height          =   195
      Left            =   4080
      TabIndex        =   27
      Top             =   3600
      Width           =   1725
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Quantidade Vendido:"
      Height          =   195
      Left            =   480
      TabIndex        =   26
      Top             =   6360
      Width           =   1725
   End
   Begin VB.Label lblQtd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   25
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "ComissãoTotal:"
      Height          =   195
      Left            =   6720
      TabIndex        =   24
      Top             =   6360
      Width           =   1320
   End
   Begin VB.Label lblVendido 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      Top             =   6360
      Width           =   1485
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Funcionário:"
      Height          =   195
      Left            =   1200
      TabIndex        =   13
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Vendido:"
      Height          =   195
      Left            =   3720
      TabIndex        =   22
      Top             =   6360
      Width           =   1365
   End
   Begin VB.Label lblComissao 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Código:"
      Height          =   195
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Produto:"
      Height          =   195
      Left            =   4080
      TabIndex        =   6
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmRelatVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Indice As String, Indice2 As String, Coluna As Integer, Coluna2 As Integer

Private Sub Imprime2()
Dim StrTemp As String, Largura As Double
Dim Qtd As Double, QtdParcial As Double
Dim VTotal As Currency, VParcial As Currency
Dim Comissao As Currency, ComissaoParcial As Currency
Dim Quebra As String, Campo As String
Dim NovaPagina As Boolean

With qVendas2
  If .Recordset.RecordCount = 0 Then
    MsgBox "Não existe registro a ser impresso!"
    Exit Sub
  End If
  
  On Error GoTo TrataErro
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  Resposta = MsgBox("Deseja salto de página a cada mudança de item indexado!", vbYesNo)
  If Resposta = vbYes Then
    NovaPagina = True
  Else
    NovaPagina = False
  End If
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  
  Largura = 190
  Dia = Now
  
  Cabeca Dia, Largura
  
  Campo = DBGrid1.Columns(Coluna).DataField
  .Recordset.MoveFirst
  Quebra = .Recordset(Campo)
  .Recordset.MoveFirst
  Do While .Recordset.EOF = False
    
    If cboProduto.Text = "" Then
      If Quebra <> .Recordset(Campo) Then
        Printer.CurrentY = Printer.CurrentY + 1
        Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
        Printer.CurrentY = Printer.CurrentY + 1
        
        StrTemp = Format(QtdParcial, "#,##0")
        Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(VParcial, "#,##0.00")
        Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp;
        
        StrTemp = Format(ComissaoParcial, "#,##0.00")
        Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
        Printer.Print StrTemp
        
        Quebra = .Recordset(Campo)
        QtdParcial = 0
        VParcial = 0
        ComissaoParcial = 0
        If NovaPagina = True Then
          Printer.CurrentY = 0
          Printer.NewPage
          Cabeca Dia, Largura
        Else
          Cabeca2 Dia, Largura
        End If
      End If
    End If
    If Printer.CurrentY > Printer.ScaleHeight - 25 Then
      Printer.CurrentY = Printer.CurrentY + 1
      Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
      Printer.CurrentY = Printer.CurrentY + 1
      
      StrTemp = Format(QtdParcial, "#,##0")
      Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(VParcial, "#,##0.00")
      Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp;
      
      StrTemp = Format(ComissaoParcial, "#,##0.00")
      Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
      Printer.Print StrTemp
      
      Printer.CurrentY = 0
      Printer.NewPage
      Cabeca Dia, Largura
    End If
    
    StrTemp = Format(.Recordset("datacaixa"), "dd/mm/yy ddd")
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Turno
    Printer.CurrentX = 20
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!CodProduto
    Printer.CurrentX = 40 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Descri
    Printer.CurrentX = 41
    Printer.Print StrTemp;
    
    Qtd = Qtd + .Recordset!Quantidade
    QtdParcial = QtdParcial + .Recordset!Quantidade
    StrTemp = .Recordset!Quantidade
    Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    
    StrTemp = Format(.Recordset!valorUnitario, "#,##0.00")
    Printer.CurrentX = 103 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    VTotal = VTotal + .Recordset("valortotal")
    VParcial = VParcial + .Recordset("valortotal")
    StrTemp = Format(.Recordset("valortotal"), "#,##0.00")
    Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    Comissao = Comissao + .Recordset!ValorComissao
    ComissaoParcial = ComissaoParcial + .Recordset!ValorComissao
    StrTemp = Format(.Recordset!ValorComissao, "#,##0.00")
    Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Codigo
    Printer.CurrentX = 138 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!Nome
    Printer.CurrentX = 139
    Printer.Print StrTemp
    
    .Recordset.MoveNext
  Loop
  Printer.CurrentY = Printer.CurrentY + 1
  Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
  Printer.CurrentY = Printer.CurrentY + 1
  
  StrTemp = Format(QtdParcial, "#,##0")
  Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(VParcial, "#,##0.00")
  Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp;
  
  StrTemp = Format(ComissaoParcial, "#,##0.00")
  Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
  Printer.Print StrTemp
  
End With
Printer.EndDoc
TrataErro:

End Sub

Private Sub Cabeca(ByVal Dia As Date, ByVal Largura As Double)
Dim StrTemp As String

Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"
Printer.FontSize = 14

StrTemp = "Relatório de Vendas"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
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

If cboProduto.Text <> "" Then
  StrTemp = "Código do Produto: " & txtCodigo.Text & "    Produto: " & cboProduto.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If
If cboFuncionario.Text <> "" Then
  StrTemp = "Código do Funcionário: " & txtCodFun.Text & "    Funcionário: " & cboFuncionario.Text
  Printer.CurrentX = 0
  Printer.Print StrTemp
End If

Printer.Print ""

StrTemp = "Dt. Venda"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Cod."
Printer.CurrentX = 40 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Produto"
Printer.CurrentX = 41
Printer.Print StrTemp;

StrTemp = "Qtd."
Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "V. Unitário"
Printer.CurrentX = 103 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total"
Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Comissão"
Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cod."
Printer.CurrentX = 138 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Funcionário"
Printer.CurrentX = 139
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub Cabeca2(ByVal Dia As Date, ByVal Largura As Double)
Dim StrTemp As String

Printer.Print ""

StrTemp = "Dt. Venda"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Turno"
Printer.CurrentX = 20
Printer.Print StrTemp;

StrTemp = "Cod."
Printer.CurrentX = 40 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Produto"
Printer.CurrentX = 41
Printer.Print StrTemp;

StrTemp = "Qtd."
Printer.CurrentX = 89 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "V. Unitário"
Printer.CurrentX = 103 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Total"
Printer.CurrentX = 115 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Comissão"
Printer.CurrentX = 129 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cod."
Printer.CurrentX = 138 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Funcionário"
Printer.CurrentX = 139
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

Private Sub cboProduto_KeyDown(KeyCode As Integer, Shift As Integer)
KeyCode = VerificaTecla(KeyCode)
End Sub

Private Sub cboProduto_LostFocus()
With dbProdutos
  .Refresh
  If cboProduto.Text = "" Then Exit Sub
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  .Recordset.Find "descri='" & cboProduto.Text & "'"
  If .Recordset.EOF = False Then
    cboProduto.Text = .Recordset!Descri
    txtCodigo.Text = .Recordset!Codigo
  End If
End With
End Sub


Private Sub cmdExibir_Click()
Dim StrTemp As String
Dim StrTemp2 As String
Dim strTemp3 As String
Dim strTemp4 As String
Dim Qtd As Double, Vendido As Currency, Comissao As Currency

Screen.MousePointer = vbHourglass

If chkNaoFinalizados.Value = vbUnchecked Then
  StrTemp = "select *from venda2 where fechamentodiario=-1 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  StrTemp2 = "select sum(quantidade) as qtd, sum(valortotal) as Vendido from venda2 where fechamentodiario=-1 and data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  
  strTemp3 = "select *from qcomissoes2 where fechamentodiario=-1 and fechamentodecaixa.datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  strTemp4 = "select sum(quantidade) as qtd, sum(valortotal) as Vendido, sum(valorcomissao) as comissao from qcomissoes2 where fechamentodiario=-1 and fechamentodecaixa.datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
Else
  StrTemp = "select *from venda2 where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  StrTemp2 = "select sum(quantidade) as qtd, sum(valortotal) as Vendido from venda2 where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  
  strTemp3 = "select *from qcomissoes2 where fechamentodecaixa.datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  strTemp4 = "select sum(quantidade) as qtd, sum(valortotal) as Vendido, sum(valorcomissao) as comissao from qcomissoes2 where fechamentodecaixa.datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
End If
If chkExibirTodas.Value = vbUnchecked Then
  If chkPago.Value = vbChecked Then
    strTemp3 = strTemp3 & " and pago=-1"
    strTemp4 = strTemp4 & " and pago=-1"
  Else
    strTemp3 = strTemp3 & " and pago=0"
    strTemp4 = strTemp4 & " and pago=0"
  End If
End If

If cboProduto.Text <> "" Then
  Call cboProduto_LostFocus
  StrTemp = StrTemp & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
  StrTemp2 = StrTemp2 & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
  strTemp3 = strTemp3 & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
  strTemp4 = strTemp4 & " and codigoproduto=" & dbProdutos.Recordset!CodigoProduto
End If
If cboFuncionario.Text <> "" Then
  Call cboFuncionario_LostFocus
  StrTemp = StrTemp & " and codigo=" & dbFuncionarios.Recordset!Codigo
  StrTemp2 = StrTemp2 & " and codigo=" & dbFuncionarios.Recordset!Codigo
  strTemp3 = strTemp3 & " and codigo=" & dbFuncionarios.Recordset!Codigo
  strTemp4 = strTemp4 & " and codigo=" & dbFuncionarios.Recordset!Codigo
End If
StrTemp = StrTemp & " order by data, codigovenda"

If Indice2 <> "" Then
  strTemp3 = strTemp3 & Indice2 & ",data, codigovenda"
End If
qVendas.RecordSource = StrTemp
qVendas.Refresh
qVendas2.RecordSource = strTemp3
qVendas2.Refresh
qVendasTotal.RecordSource = StrTemp2
qVendasTotal.Refresh

Qtd = 0
Vendido = 0
Comissao = 0

With qVendasTotal
  If IsNull(.Recordset!Qtd) = False Then
    lblQtd1.Caption = .Recordset!Qtd
  Else
    lblQtd1.Caption = "0"
  End If
  If IsNull(.Recordset!Vendido) = False Then
    lblTotalVendido1.Caption = Format(.Recordset!Vendido, "Currency")
  Else
    lblTotalVendido1.Caption = Format("0", "Currency")
  End If
End With
With qVendasTotal
  .RecordSource = strTemp4
  .Refresh
  .Refresh
  If IsNull(.Recordset!Qtd) = False Then
    Qtd = Qtd + .Recordset!Qtd
  End If
  If IsNull(.Recordset!Vendido) = False Then
    Vendido = Vendido + .Recordset!Vendido
  End If
  If IsNull(.Recordset!Comissao) = False Then
    Comissao = Comissao + .Recordset!Comissao
  End If
End With
lblQtd.Caption = Qtd
lblVendido.Caption = Format(Vendido, "Currency")
lblComissao.Caption = Format(Comissao, "Currency")

Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double
Dim Qtd As Double, QtdParcial As Double
Dim VTotal As Currency, VParcial As Currency
Dim Comissao As Currency, ComissaoParcial As Currency
Dim Quebra As Integer
Dim NovaPagina As Boolean


Resposta = MsgBox("Deseja imprimir a venda de produtos comissionados e não comissionados?", vbYesNo)
If Resposta = vbYes Then
  With qVendas
    If .Recordset.RecordCount = 0 Then
      MsgBox "Não existe registro a ser impresso!"
      GoTo TrataErro
    End If
    
    On Error GoTo TrataErro
    If ShowPrinter(Me) = 0 Then GoTo TrataErro
    On Error GoTo 0
    
    Resposta = MsgBox("Deseja salto de página a cada mudança de item indexado!", vbYesNo)
    If Resposta = vbYes Then
      Quebra = Coluna2
    Else
      Quebra = 0
    End If
  End With
  
  StrTemp = StrTemp & Chr(vbKeyReturn) & "Data: " & Format(Dia, "long date")
  StrTemp = StrTemp & Chr(vbKeyReturn) & "Página: " & Printer.Page
  StrTemp = StrTemp & Chr(vbKeyReturn) & "Período: " & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "Short date")
  If cboProduto.Text <> "" Then
    StrTemp = StrTemp & Chr(vbKeyReturn) & "Código do Produto: " & txtCodigo.Text & "    Produto: " & cboProduto.Text
  End If
  If cboFuncionario.Text <> "" Then
    StrTemp = StrTemp & Chr(vbKeyReturn) & "Código do Funcionário: " & txtCodFun.Text & "    Funcionário: " & cboFuncionario.Text
  End If
  

  ImprimeADOGrid DataGrid1, Printer, qVendas, 3, True, , Quebra, 5, , "Relatório de Vendas" & Chr(vbKeyReturn) & NomePosto, "Impresso em: " & Format(Date, "Long Date"), StrTemp
  Printer.EndDoc
End If

TrataErro:
Imprime2
End Sub

Private Sub cmdPagar_Click()
Dim Resposta As Integer, Total As Currency
Total = 0
If chkPago.Value = vbChecked Then
  MsgBox "Comissões já pagas, não podem ser pagas novamente!"
  Exit Sub
End If
If chkExibirTodas.Value = vbChecked Then
  MsgBox "Pode haver a ocorrência de comissoes já pagas!" & Chr(13) & "Faça nova exibição retirando a opção de Todas as Vendas!"
  Exit Sub
End If

Total = -CCur(lblComissao.Caption)
Resposta = MsgBox("Deseja pagar as comissões exibidas na tela?", vbYesNo)
If Resposta = vbNo Then Exit Sub
With qVendas2
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbVendas2.Recordset.FindFirst "codigovenda=" & .Recordset!codigovenda
      dbVendas2.Recordset.Edit
      dbVendas2.Recordset!Pago = True
      dbVendas2.Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
End With
With dbDespesaLanc
  .Recordset.AddNew
  .Recordset!CodigoFechamento = 0
  .Recordset!Origem = "Despesa"
  .Recordset!Data = Date
  .Recordset!Hora = Now
  .Recordset!Vencimento = Date
  .Recordset!CodigoConta = 0
  .Recordset!CodigoDespesa = 0
  .Recordset!Descri = "Comissões-" & Format(txtDataIni.Value, "short date") & " a " & Format(txtDataFim.Value, "short date")
  .Recordset!obs = cboFuncionario.Text & " | " & cboProduto.Text
  .Recordset!Valor = Total
  .Recordset!Fechamento = True
  .Recordset.Update
End With
Call cmdExibir_Click
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
Coluna2 = ColIndex
End Sub

Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
If Indice2 = " order by " & DBGrid1.Columns(ColIndex).DataField Then
  Indice2 = " order by " & DBGrid1.Columns(ColIndex).DataField & " desc"
Else
  Indice2 = " order by " & DBGrid1.Columns(ColIndex).DataField
End If
Coluna = ColIndex
StrTemp = qVendas.Recordset.Sort
Call cmdExibir_Click
qVendas.Recordset.Sort = StrTemp
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
KeyAscii = VerificaTecla(KeyAscii)
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With qVendas
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from qvendas where codigovenda=0 order by venda2.data, codigovenda"
  .Refresh
End With
With qVendasTotal
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(venda2.valortotal) as Vendido, sum(valorcomissao) as comissao from qvendas where codigovenda=0"
  .Refresh
End With
With dbFuncionarios
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from vendedores order by nome"
  .Refresh
End With
With dbProdutos
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from produtos where combustivel=0 order by descri"
  .Refresh
End With
With dbDespesaLanc
  .ConnectionString = CaminhoADO
  .Refresh
End With
With dbVendas2
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With qVendas2
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
Select Case Usuarios.Grupo.RelatVendaProdutos
  Case 1 'Somente leitura
    cmdPagar.Enabled = False
  Case 2 'Liberado
    
End Select

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

Private Sub txtCodigo_GotFocus()
With txtCodigo
  .SelStart = 0
  .SelLength = Len(.Text)
End With
End Sub

Private Sub txtCodigo_LostFocus()
With txtCodigo
  If .Text = "" Then Exit Sub
  dbProdutos.Refresh
  If dbProdutos.Recordset.RecordCount = 0 Then Exit Sub
  dbProdutos.Recordset.MoveFirst
  dbProdutos.Recordset.Find "codigo=" & .Text
  If dbProdutos.Recordset.EOF = False Then
    cboProduto.Text = dbProdutos.Recordset!Descri
    txtCodigo.Text = dbProdutos.Recordset!Codigo
  End If
End With
End Sub

Private Sub txtDataFim_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case vbKeyReturn
    txtCodigo.SetFocus
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
