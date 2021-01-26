VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVendasLeituraX 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vendas e Leitura X"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
   Icon            =   "frmVendasLeituraX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc dbConfig 
      Height          =   375
      Left            =   4920
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbConfig"
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
   Begin MSAdodcLib.Adodc dbImportacao 
      Height          =   375
      Left            =   4920
      Top             =   3000
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "dbImportacao"
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
   Begin VB.CommandButton cmdImportar 
      Caption         =   "Importar"
      Height          =   255
      Left            =   8760
      TabIndex        =   32
      Top             =   7560
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   360
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   11520
      Picture         =   "frmVendasLeituraX.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   11
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin VB.ComboBox cboCombustivel 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc qVendasLeituraX 
      Height          =   375
      Left            =   4920
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      RecordSource    =   "select *from vendasleiturax"
      Caption         =   "qVendasLeituraX"
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
   Begin MSAdodcLib.Adodc dbVendasLeituraX 
      Height          =   375
      Left            =   4920
      Top             =   2040
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
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
      RecordSource    =   "select *from vendasleiturax"
      Caption         =   "dbVendasLeituraX"
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
      Bindings        =   "frmVendasLeituraX.frx":0EC4
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   10186
      _Version        =   393216
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
         DataField       =   "Data"
         Caption         =   "Data"
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
         DataField       =   "Categoria"
         Caption         =   "Categoria"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "d/M/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "SistemaQtd"
         Caption         =   "Sistema Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " #,##0.#0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "SistemaValor"
         Caption         =   "Sistema Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "currency"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "LeituraXQtd"
         Caption         =   "Leitura X Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " #,##0.#0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "LeituraXValor"
         Caption         =   "Leitura X Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "currency"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "PrecoDiferenciado"
         Caption         =   "Diferenciado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "currency"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "DiferencaQtd"
         Caption         =   "Dif. Qtd"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   " #,##0.#0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "DiferencaValor"
         Caption         =   "Dif. Valor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "currency"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "ReducaoZ"
         Caption         =   "Redução Z"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "currency"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         BeginProperty Column00 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   2369,764
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1049,953
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1140,095
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   989,858
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   840,189
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   929,764
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
            ColumnWidth     =   1124,787
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1095
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
      CurrentDate     =   37678
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
      CurrentDate     =   37678
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Diferença Redução Z:"
      Height          =   195
      Left            =   10680
      TabIndex        =   30
      Top             =   6720
      Width           =   1590
   End
   Begin VB.Label lblDifReducaoZ 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   10680
      TabIndex        =   29
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label lblReducaoZ 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9120
      TabIndex        =   28
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Redução Z:"
      Height          =   195
      Left            =   9120
      TabIndex        =   27
      Top             =   6720
      Width           =   855
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Left            =   5640
      TabIndex        =   26
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "="
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "-"
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   6960
      Width           =   255
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Diferenciado:"
      Height          =   195
      Left            =   4080
      TabIndex        =   21
      Top             =   7320
      Width           =   945
   End
   Begin VB.Label lblDiferenciado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   20
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Sistema Qtd.:"
      Height          =   195
      Left            =   2160
      TabIndex        =   19
      Top             =   6720
      Width           =   945
   End
   Begin VB.Label lblSistemaQtd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   18
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Sistema Valor:"
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   7320
      Width           =   1005
   End
   Begin VB.Label lblSistemaValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2160
      TabIndex        =   16
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Leitura X Qtd.:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   6720
      Width           =   1020
   End
   Begin VB.Label lblLeituraXQtd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Leitura X Valor:"
      Height          =   195
      Left            =   240
      TabIndex        =   13
      Top             =   7320
      Width           =   1080
   End
   Begin VB.Label lblLeituraXValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label lblTotalValor 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   7560
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Diferença Valor:"
      Height          =   195
      Left            =   6000
      TabIndex        =   9
      Top             =   7320
      Width           =   1140
   End
   Begin VB.Label lblTotalQtd 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Diferença Qtd.:"
      Height          =   195
      Left            =   4080
      TabIndex        =   7
      Top             =   6720
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   90
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
End
Attribute VB_Name = "frmVendasLeituraX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AtualizaSaldo()
Dim StrCombustivel As String
Select Case cboCombustivel.Text
  Case "Todos"
    StrCombustivel = ""
  Case Else
    StrCombustivel = " and categoria='" & cboCombustivel.Text & "'"
End Select
With qVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "select sum(diferencaqtd) as qtd, sum(diferencavalor) as total, sum(sistemaqtd) as sisqtd, sum(sistemavalor) as sisvalor, sum(leituraxqtd) as leituraqtd, sum(leituraxvalor) as leituravalor, sum(precodiferenciado) as diferenciado, sum(reducaoz) as Z from vendasleiturax where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrCombustivel
  .Refresh
  If IsNull(.Recordset!Qtd) = False Then
    lblTotalQtd.Caption = Format(.Recordset!Qtd, "0.00")
  Else
    lblTotalQtd.Caption = Format(0, "0.00")
  End If
  If IsNull(.Recordset!Total) = False Then
    lblTotalValor.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblTotalValor.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!sisvalor) = False Then
    lblSistemaValor.Caption = Format(.Recordset!sisvalor, "Currency")
  Else
    lblSistemaValor.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!sisqtd) = False Then
    lblSistemaQtd.Caption = Format(.Recordset!sisqtd, "0.00")
  Else
    lblSistemaQtd.Caption = Format(0, "0.00")
  End If
  If IsNull(.Recordset!leituraqtd) = False Then
    lblLeituraXQtd.Caption = Format(.Recordset!leituraqtd, "0.00")
  Else
    lblLeituraXQtd.Caption = Format(0, "0.00")
  End If
  If IsNull(.Recordset!leituravalor) = False Then
    lblLeituraXValor.Caption = Format(.Recordset!leituravalor, "Currency")
  Else
    lblLeituraXValor.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!diferenciado) = False Then
    lblDiferenciado.Caption = Format(.Recordset!diferenciado, "Currency")
  Else
    lblDiferenciado.Caption = Format(0, "Currency")
  End If
  
  
  If IsNull(.Recordset!z) = False Then
    lblReducaoZ.Caption = Format(.Recordset!z, "Currency")
  Else
    lblReducaoZ.Caption = Format(0, "Currency")
  End If
  If IsNull(.Recordset!z) = False And IsNull(.Recordset!sisvalor) = False And IsNull(.Recordset!diferenciado) = False Then
    lblDifReducaoZ.Caption = Format(.Recordset!z - .Recordset!sisvalor - .Recordset!diferenciado, "Currency")
  Else
    lblDifReducaoZ.Caption = Format(0, "Currency")
  End If
  
End With
End Sub

Private Sub cmdExibir_Click()
Dim StrCombustivel As String
Dim db As New ADODB.Connection
Dim dbVendas As New ADODB.Recordset
Dim dbLeituraX As New ADODB.Recordset
Dim dbEncerrantes As New ADODB.Recordset
Dim dbNotas As New ADODB.Recordset
Dim dbTemp As New ADODB.Recordset
Dim dbGrupoIf As New ADODB.Recordset
Dim dbPDVs As New ADODB.Recordset

Dim Resposta As Integer

Dim CodigoPdv As Double



db.Open CaminhoADO
Resposta = MsgBox("Deseja Apagar os registros gerados?", vbYesNo + vbDefaultButton2)

If Resposta = vbYes Then
  Select Case cboCombustivel.Text
    Case "Todos"
      StrCombustivel = ""
    Case Else
      StrCombustivel = " and departamento='" & cboCombustivel.Text & "'"
  End Select
  
  Screen.MousePointer = vbHourglass
  
  db.Execute "delete from vendasleiturax where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  
End If

Resposta = MsgBox("Deseja calcular agora?", vbYesNo)

If Resposta = vbYes Then
  Select Case cboCombustivel.Text
    Case "Todos"
      StrCombustivel = ""
    Case Else
      StrCombustivel = " and departamento='" & cboCombustivel.Text & "'"
  End Select
  
  Screen.MousePointer = vbHourglass
  cmdExibir.Enabled = False
  DataGrid1.Visible = False
  ProgressBar1.Visible = True
  
  dbTemp.CursorLocation = adUseServer
  dbTemp.Open "select *from vendasleiturax where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#", db, adOpenKeyset, adLockOptimistic
  If dbTemp.RecordCount <> 0 Then
    dbTemp.MoveLast
    dbTemp.MoveFirst
    Do While dbTemp.EOF = False
      If dbTemp!LeituraXQtd = 0 And dbTemp!LeituraXValor = 0 And dbTemp!reducaoz = 0 Then
        dbTemp.Delete adAffectCurrent
      End If
      If IsNull(dbTemp!LeituraXQtd) = True And IsNull(dbTemp!LeituraXValor) = True And IsNull(dbTemp!reducaoz) = True Then
        dbTemp.Delete adAffectCurrent
      End If
      If IsNull(dbTemp!LeituraXQtd) = True Then
        dbTemp!LeituraXQtd = 0
        dbTemp.Update
      End If
      If IsNull(dbTemp!LeituraXValor) = True Then
        dbTemp!LeituraXValor = 0
        dbTemp.Update
      End If
      If IsNull(dbTemp!reducaoz) = True Then
        dbTemp!reducaoz = 0
        dbTemp.Update
      End If
      
      dbTemp.MoveNext
    Loop
  End If
  dbTemp.Close
  
  dbPDVs.Open "select *from pdvs", db, adOpenKeyset, adLockOptimistic
  
  If dbPDVs.RecordCount <> 0 Then
      CodigoPdv = dbPDVs!CodigoPdv
  End If
  dbVendas.Open "select departamento, data, sum(vendido) as qtd, sum(valor) as total, codigogrupoif from qvendadiaprodutos2 where combustivel=0 and data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# group by departamento, data, codigogrupoif", db, adOpenKeyset, adLockOptimistic
  dbEncerrantes.Open "select datacaixa, departamento, sum(Encerrante - Abertura) as vendas, sum (retorno) as ret, sum(valortotal) as total, codigogrupoif from qbicoencerrantes where codigopdv=" & CodigoPdv & " and datacaixa between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# group by departamento, datacaixa, codigogrupoif", db, adOpenKeyset, adLockOptimistic
  dbNotas.Open "select produtos.departamento, produtos.combustivel, clientesnota2.data, sum(lucrodif) as total, produtos.codigogrupoif from clientesnota2, produtos where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# and produtos.codigo=clientesnota2.codigoproduto group by departamento, combustivel, data, produtos.codigogrupoif", db, adOpenKeyset, adLockOptimistic
  dbGrupoIf.CursorLocation = adUseClient
  dbGrupoIf.Open "Select codigo, codigogrupo, descri from produtosgrupoif", db, adOpenKeyset, adLockOptimistic
  
  
  'dbLeituraX.Open "select *from vendasleiturax where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "# order by data, categoria", Db, adOpenKeyset, adLockOptimistic
  On Error Resume Next
  db.Execute "update vendasleiturax set sistemaqtd=0 where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrCombustivel
  db.Execute "update vendasleiturax set sistemavalor=0 where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrCombustivel
  db.Execute "update vendasleiturax set precodiferenciado=0 where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrCombustivel
  On Error GoTo 0
  
  If dbVendas.EOF = False And dbVendas.BOF = False Then
    On Error Resume Next
    dbVendas.MoveLast
    ProgressBar1.Max = dbVendas.RecordCount
    dbVendas.MoveFirst
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    
    
    Do While dbVendas.EOF = False
      StrTemp = dbVendas!departamento
      If dbGrupoIf.RecordCount <> 0 Then
        dbGrupoIf.MoveFirst
        If IsNull(dbVendas!codigogrupoif) = False Then
          dbGrupoIf.Find "codigo=" & dbVendas!codigogrupoif
          If dbGrupoIf.EOF = False Then
            StrTemp = dbGrupoIf!CodigoGrupo & " " & dbGrupoIf!Descri
            'StrTemp = dbGrupoIf!Descri
            
          End If
        End If
      End If
      dbLeituraX.Open "select *from vendasleiturax where data=#" & DataInglesa(dbVendas!Data) & "# and categoria='" & StrTemp & "'", db, adOpenKeyset, adLockOptimistic
      If dbLeituraX.EOF = True And dbLeituraX.BOF = True Then
        dbLeituraX.AddNew
        dbLeituraX!Data = dbVendas!Data
        dbLeituraX!Categoria = StrTemp
        dbLeituraX!codigogrupoif = dbVendas!codigogrupoif
        dbLeituraX!Combustivel = 0
        dbLeituraX!sistemaqtd = 0
        dbLeituraX!sistemavalor = 0
        dbLeituraX!LeituraXQtd = 0
        dbLeituraX!LeituraXValor = 0
        dbLeituraX.Update
      End If
      dbLeituraX.Close
      On Error Resume Next
      db.Execute "update vendasleiturax set sistemaqtd=" & Replace(dbVendas!Qtd, ",", ".") & " where data=#" & DataInglesa(dbVendas!Data) & "# and categoria='" & StrTemp & "'"
      db.Execute "update vendasleiturax set sistemavalor=" & Replace(dbVendas!Total, ",", ".") & " where data=#" & DataInglesa(dbVendas!Data) & "# and categoria='" & StrTemp & "'"
      db.Execute "update vendasleiturax set combustivel=0 where data=#" & DataInglesa(dbVendas!Data) & "# and categoria='" & StrTemp & "'"
      dbVendas.MoveNext
      ProgressBar1.Value = dbVendas.AbsolutePosition
      DoEvents
      On Error GoTo 0
    Loop
  End If
  
  
  If dbEncerrantes.EOF = False And dbEncerrantes.BOF = False Then
    dbEncerrantes.MoveLast
    dbEncerrantes.MoveFirst
    On Error Resume Next
    ProgressBar1.Max = dbEncerrantes.RecordCount
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    On Error GoTo 0
    Do While dbEncerrantes.EOF = False
      StrTemp = dbEncerrantes!departamento
      If dbGrupoIf.RecordCount <> 0 Then
        If IsNull(dbEncerrantes!codigogrupoif) = False Then
          dbGrupoIf.MoveFirst
          dbGrupoIf.Find "codigo=" & dbEncerrantes!codigogrupoif
          If dbGrupoIf.EOF = False Then
            StrTemp = dbGrupoIf!CodigoGrupo & " " & dbGrupoIf!Descri
            'StrTemp = dbGrupoIf!Descri
          End If
        End If
      End If
      dbLeituraX.Open "select *from vendasleiturax where data=#" & DataInglesa(dbEncerrantes!DataCaixa) & "# and categoria='" & StrTemp & "'", db, adOpenKeyset, adLockOptimistic
      If dbLeituraX.EOF = True And dbLeituraX.BOF = True Then
        dbLeituraX.AddNew
        dbLeituraX!Data = dbEncerrantes!DataCaixa
        dbLeituraX!Categoria = StrTemp
        dbLeituraX!Combustivel = -1
        dbLeituraX!sistemaqtd = 0
        dbLeituraX!sistemavalor = 0
        dbLeituraX!LeituraXQtd = 0
        dbLeituraX!LeituraXValor = 0
        dbLeituraX.Update
      End If
      dbLeituraX.Close
      Vendido = (dbEncerrantes!Vendas - dbEncerrantes!Ret)
      On Error Resume Next
      db.Execute "update vendasleiturax set sistemaqtd=" & Replace(Vendido, ",", ".") & " where data=#" & DataInglesa(dbEncerrantes!DataCaixa) & "# and categoria='" & StrTemp & "'"
      db.Execute "update vendasleiturax set sistemavalor=" & Replace(dbEncerrantes!Total, ",", ".") & " where data=#" & DataInglesa(dbEncerrantes!DataCaixa) & "# and categoria='" & StrTemp & "'"
      db.Execute "update vendasleiturax set combustivel=-1 where data=#" & DataInglesa(dbEncerrantes!DataCaixa) & "# and categoria='" & StrTemp & "'"
      dbEncerrantes.MoveNext
      
      ProgressBar1.Value = dbEncerrantes.AbsolutePosition
      DoEvents
      On Error GoTo 0
    Loop
  End If
  If dbNotas.EOF = False And dbNotas.BOF = False Then
    dbNotas.MoveLast
    dbNotas.MoveFirst
    On Error Resume Next
    ProgressBar1.Max = dbNotas.RecordCount
    ProgressBar1.Value = 0
    ProgressBar1.Visible = True
    On Error GoTo 0
    Do While dbNotas.EOF = False
      StrTemp = dbNotas!departamento
      If dbGrupoIf.RecordCount <> 0 Then
        dbGrupoIf.MoveFirst
        If IsNull(dbNotas!codigogrupoif) = False Then
          dbGrupoIf.Find "codigo=" & dbNotas!codigogrupoif
          If dbGrupoIf.EOF = False Then
            StrTemp = dbGrupoIf!CodigoGrupo & " " & dbGrupoIf!Descri
          End If
        End If
      End If
      dbLeituraX.Open "select *from vendasleiturax where data=#" & DataInglesa(dbNotas!Data) & "# and categoria='" & StrTemp & "'", db, adOpenKeyset, adLockOptimistic
      If dbLeituraX.EOF = True And dbLeituraX.BOF = True Then
        dbLeituraX.AddNew
        dbLeituraX!Data = dbNotas!Data
        dbLeituraX!Categoria = StrTemp
        dbLeituraX!Combustivel = dbNotas!Combustivel
        dbLeituraX!sistemaqtd = 0
        dbLeituraX!sistemavalor = 0
        dbLeituraX!LeituraXQtd = 0
        dbLeituraX!LeituraXValor = 0
        dbLeituraX.Update
      End If
      dbLeituraX.Close
      On Error Resume Next
      db.Execute "update vendasleiturax set precodiferenciado=" & Replace(dbNotas!Total, ",", ".") & " where data=#" & DataInglesa(dbNotas!Data) & "# and categoria='" & dbNotas!departamento & "'"
      dbNotas.MoveNext
      
      ProgressBar1.Value = dbNotas.AbsolutePosition
      DoEvents
      On Error GoTo 0
    Loop
  End If
  
  On Error Resume Next
  db.Execute "update vendasleiturax set diferencaqtd=leituraxqtd-sistemaqtd where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  db.Execute "update vendasleiturax set diferencavalor=leituraxvalor-sistemavalor-precodiferenciado where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#"
  On Error GoTo 0
  dbVendas.Close
  dbEncerrantes.Close
  dbGrupoIf.Close

End If

db.Execute "update vendasleiturax set LeituraXValor=0 where LeituraXValor=null"
db.Execute "update vendasleiturax set LeituraXQtd=0 where LeituraXQtd=null"
db.Execute "update vendasleiturax set PrecoDiferenciado=0 where PrecoDiferenciado=null"
db.Execute "update vendasleiturax set ReducaoZ=0 where ReducaoZ=null"
db.Close

Select Case cboCombustivel.Text
  Case "Todos"
    StrCombustivel = ""
  Case Else
    StrCombustivel = " and categoria='" & cboCombustivel.Text & "'"
End Select

With dbVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from vendasleiturax where data between #" & DataInglesa(txtDataIni.Value) & "# and #" & DataInglesa(txtDataFim.Value) & "#" & StrCombustivel & " order by data, categoria"
  .Refresh
End With

AtualizaSaldo
ProgressBar1.Visible = False
Screen.MousePointer = vbDefault
cmdExibir.Enabled = True
DataGrid1.Visible = True
End Sub

Private Sub cmdImportar_Click()
Dim Dia As Date, strEncerrantes As String, intArquivo As Integer
Dim StrTemp As String, SoPrimeira As Boolean

cmdImportar.Enabled = False
With dbConfig
  .ConnectionString = CaminhoADO
  .RecordSource = "select *from config"
  .Refresh
End With

With dbImportacao
  .ConnectionString = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbConfig.Recordset!ftp
  .RecordSource = "select *from caixas where linhaexportada like '007%' and datacaixa between '" & txtDataIni.Value & "' and '" & txtDataFim.Value & "' and codigoposto='" & dbConfig.Recordset!Porta & "' order by linhaexportada"
  .Refresh
  If .Recordset.RecordCount = 0 Then
    MsgBox "O caixa atual ainda não foi exportado!"
    cmdImportar.Enabled = True
    Exit Sub
  End If
  .Recordset.MoveLast
  .Recordset.MoveFirst
  
  Do While .Recordset.EOF = False
    StrTemp = .Recordset!linhaexportada
    DoEvents
    Select Case Mid(StrTemp, 1, 3)
      Case "007"
        GravaCupons StrTemp, dbVendasLeituraX
    End Select
    .Recordset.MoveNext
  Loop
End With

'AtualizaSaldo

dbVendasLeituraX.Refresh

MsgBox "Importação finalizada!"

cmdImportar.Enabled = True
End Sub

Private Sub cmdImprime_Click()
On Error GoTo NaoImprime
If ShowPrinter(Me) = 0 Then GoTo NaoImprime
On Error GoTo 0

ImprimeADOGrid DataGrid1, Printer, dbVendasLeituraX, 2, , , 0, 3, 4, NomePosto, "Leitura X e Vendas", "Período: " & txtDataIni.Value & " a " & txtDataFim.Value & " - Categorias:" & cboCombustivel.Text, 5, 6, 7, 8

Printer.CurrentX = 0
Printer.Print "Vendas no sistema:" & lblSistemaValor.Caption
Printer.CurrentX = 0
Printer.Print "Total da Redução Z:" & lblReducaoZ.Caption
Printer.CurrentX = 0
Printer.Print "Diferença da Redução Z:" & lblDifReducaoZ.Caption



Printer.EndDoc
NaoImprime:

End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
With dbVendasLeituraX
  .Recordset!diferencaqtd = .Recordset!LeituraXQtd - .Recordset!sistemaqtd
  .Recordset!diferencavalor = .Recordset!LeituraXValor - .Recordset!sistemavalor
  On Error Resume Next
  .Recordset.Update
End With
AtualizaSaldo
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
Dim dbProdutos As New ADODB.Recordset

txtDataIni.Value = Date
txtDataFim.Value = Date

db.Open CaminhoADO
db.Execute "update Produtos, produtosgrupoif set produtos.departamento=produtosgrupoif.descri where produtos.codigogrupoif=produtosgrupoif.codigo"
db.Execute "update VendasLeituraX, produtosgrupoif set VendasLeituraX.categoria=produtosgrupoif.descri where VendasLeituraX.codigogrupoif=produtosgrupoif.codigo"

dbProdutos.Open "select departamento from produtos group by departamento order by departamento", db, adOpenKeyset, adLockOptimistic
cboCombustivel.Clear
If dbProdutos.RecordCount <> 0 Then
  dbProdutos.MoveLast
  dbProdutos.MoveFirst
  Do While dbProdutos.EOF = False
    If IsNull(dbProdutos!departamento) = False Then
      cboCombustivel.AddItem dbProdutos!departamento
    End If
    dbProdutos.MoveNext
  Loop
End If
cboCombustivel.AddItem "Todos"
cboCombustivel.Text = "Todos"

dbProdutos.Close
db.Close

With dbVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from vendasleiturax where categoria='desativadosemvendas'"
  .Refresh
End With
With qVendasLeituraX
  .ConnectionString = CaminhoADO
  .RecordSource = "Select *from vendasleiturax where categoria='desativadosemvendas'"
  .Refresh
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
