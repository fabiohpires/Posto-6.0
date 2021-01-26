VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRelatChequesClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatório de Cheques por Clientes"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmRelatChequesClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   4320
      TabIndex        =   19
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdImprime 
      Height          =   615
      Left            =   6720
      Picture         =   "frmRelatChequesClientes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   18
      Tag             =   "Imprimir"
      Top             =   120
      Width           =   735
   End
   Begin MSAdodcLib.Adodc dbPre 
      Height          =   330
      Left            =   960
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
      RecordSource    =   "select count(codigocheque) as Cheques, sum(valor) as Total from cheques"
      Caption         =   "dbPre"
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
   Begin MSAdodcLib.Adodc dbCobrando 
      Height          =   330
      Left            =   960
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
      RecordSource    =   "select count(codigocheque) as Cheques, sum(valor) as Total from cheques"
      Caption         =   "dbCobrando"
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
   Begin MSAdodcLib.Adodc dbDevolvidos 
      Height          =   330
      Left            =   960
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
      RecordSource    =   "select count(codigocheque) as Cheques, sum(valor) as Total from cheques"
      Caption         =   "dbDevolvidos"
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
      Left            =   3000
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin MSAdodcLib.Adodc dbDepositados 
      Height          =   330
      Left            =   960
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
      RecordSource    =   "select count(codigocheque) as Cheques, sum(valor) as Total from cheques"
      Caption         =   "dbDepositados"
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
   Begin MSAdodcLib.Adodc dbContagem 
      Height          =   330
      Left            =   960
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
      RecordSource    =   "select *from qcheques order by debanco, deagencia, deconta"
      Caption         =   "dbContagem"
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
      Bindings        =   "frmRelatChequesClientes.frx":0EC4
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5530
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "DeBanco"
         Caption         =   "Banco"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "DeAgencia"
         Caption         =   "Agência"
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
         DataField       =   "DeConta"
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
      BeginProperty Column03 
         DataField       =   "Cheques"
         Caption         =   "Cheques"
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
         DataField       =   "total"
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
            Alignment       =   1
            ColumnWidth     =   870,236
         EndProperty
         BeginProperty Column01 
            Alignment       =   1
            ColumnWidth     =   1005,165
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            ColumnWidth     =   1604,976
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1019,906
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1890,142
         EndProperty
      EndProperty
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
      Format          =   58851329
      CurrentDate     =   37678
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   300
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   529
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   37678
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Pré datados::"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   4080
      Width           =   945
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "dbPre"
      Height          =   255
      Left            =   1920
      TabIndex        =   16
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Cheques"
      DataSource      =   "dbPre"
      Height          =   255
      Left            =   1200
      TabIndex        =   15
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Cheques"
      DataSource      =   "dbCobrando"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Cheques"
      DataSource      =   "dbDevolvidos"
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Cheques"
      DataSource      =   "dbDepositados"
      Height          =   255
      Left            =   1200
      TabIndex        =   12
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "dbCobrando"
      Height          =   255
      Left            =   5640
      TabIndex        =   11
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "dbDevolvidos"
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "dbDepositados"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cobrança:"
      Height          =   195
      Left            =   4065
      TabIndex        =   8
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Devolvidos:"
      Height          =   195
      Left            =   3960
      TabIndex        =   7
      Top             =   4080
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Depositados:"
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
End
Attribute VB_Name = "frmRelatChequesClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cabeca(ByVal Largura As Double, Dia As Date)
Dim StrTemp As String
Printer.ScaleMode = vbMillimeters
Printer.FontName = "Arial"

Printer.FontSize = 14
StrTemp = "Relatório de Cheques por Cliente"
Printer.CurrentX = (Largura / 2) - (Printer.TextWidth(StrTemp) / 2)
Printer.Print StrTemp

Printer.FontSize = 8
StrTemp = "Impresso em: " & Format(Dia, "Long Date")
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Página: " & Printer.Page
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Período: " & Format(txtDataIni.Value, "Short date") & " a " & Format(txtDataFim.Value, "Short date")
Printer.CurrentX = 0
Printer.Print StrTemp

Printer.CurrentY = 13
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

StrTemp = "Não"
Printer.CurrentX = 44 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

StrTemp = "Bco"
Printer.CurrentX = 0
Printer.Print StrTemp;

StrTemp = "Ag."
Printer.CurrentX = 15 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Conta"
Printer.CurrentX = 34 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Dep."
Printer.CurrentX = 44 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Não Dep."
Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Dep."
Printer.CurrentX = 74 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Dep."
Printer.CurrentX = 94 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Dev."
Printer.CurrentX = 104 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Devol."
Printer.CurrentX = 124 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Cobr."
Printer.CurrentX = 134 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Cobr."
Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "Tot."
Printer.CurrentX = 164 - Printer.TextWidth(StrTemp)
Printer.Print StrTemp;

StrTemp = "$ Total"
Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
Printer.Print StrTemp

Printer.CurrentY = Printer.CurrentY + 1
Printer.Line (0, Printer.CurrentY)-(Largura, Printer.CurrentY)
Printer.CurrentY = Printer.CurrentY + 1

End Sub

Private Sub cmdExibir_Click()
With dbContagem
  StrTemp = .Recordset.Sort
  .ConnectionString = CaminhoADO
  .RecordSource = "SELECT Count(Cheques.CodigoFechamento) AS Cheques, Sum(Cheques.Valor) AS total, First(Cheques.Banco) AS DeBanco, First(Cheques.Agencia) AS DeAgencia, First(Cheques.Conta) AS DeConta From Cheques where cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# GROUP BY cheques.banco, cheques.agencia, cheques.conta"
  .Refresh
  .Recordset.Sort = "debanco, deagencia, deconta"
End With
End Sub

Private Sub cmdImprime_Click()
Dim StrTemp As String, Largura As Double, Dia As Date
With dbContagem
  If .Recordset.RecordCount = 0 Then Exit Sub
  .Recordset.MoveFirst
  Largura = 190
  Dia = Now
  
  On Error GoTo naoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  
  
  
  Printer.ScaleMode = vbMillimeters
  Printer.FontName = "Arial"
  Printer.DrawWidth = 2
  
  Cabeca Largura, Dia
  
  Do While .Recordset.EOF = False
    If Printer.CurrentY > Printer.ScaleHeight - 35 Then
      Y2 = Printer.CurrentY
      Printer.Line (0, Y2)-(Largura, Y2)
      
      Printer.Line (0, 13)-(0, Y2)
      Printer.Line (6, 13)-(6, Y2)
      Printer.Line (16, 13)-(16, Y2)
      Printer.Line (35, 13)-(35, Y2)
      Printer.Line (45, 13)-(45, Y2)
      Printer.Line (65, 13)-(65, Y2)
      Printer.Line (75, 13)-(75, Y2)
      Printer.Line (95, 13)-(95, Y2)
      Printer.Line (105, 13)-(105, Y2)
      Printer.Line (125, 13)-(125, Y2)
      Printer.Line (135, 13)-(135, Y2)
      Printer.Line (155, 13)-(155, Y2)
      Printer.Line (165, 13)-(165, Y2)
      Printer.Line (Largura, 13)-(Largura, Y2)
            
      Printer.NewPage
      Cabeca Largura, Dia
    End If
    DoEvents
    
    StrTemp = .Recordset!debanco
    Printer.CurrentX = 0
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!deAgencia
    Printer.CurrentX = 15 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!deconta
    Printer.CurrentX = 34 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = dbPre.Recordset!cheques
    Printer.CurrentX = 44 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(dbPre.Recordset!Total, "Currency")
    Printer.CurrentX = 64 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = dbDepositados.Recordset!cheques
    Printer.CurrentX = 74 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(dbDepositados.Recordset!Total, "Currency")
    Printer.CurrentX = 94 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = dbDevolvidos.Recordset!cheques
    Printer.CurrentX = 104 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(dbDevolvidos.Recordset!Total, "Currency")
    Printer.CurrentX = 124 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = dbCobrando.Recordset!cheques
    Printer.CurrentX = 134 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(dbCobrando.Recordset!Total, "Currency")
    Printer.CurrentX = 154 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = .Recordset!cheques
    Printer.CurrentX = 164 - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp;
    
    StrTemp = Format(.Recordset!Total, "Currency")
    Printer.CurrentX = Largura - Printer.TextWidth(StrTemp)
    Printer.Print StrTemp
    

    .Recordset.MoveNext
  Loop
  Y2 = Printer.CurrentY
  Printer.Line (0, Y2)-(Largura, Y2)
  
  Printer.Line (0, 13)-(0, Y2)
  Printer.Line (6, 13)-(6, Y2)
  Printer.Line (16, 13)-(16, Y2)
  Printer.Line (35, 13)-(35, Y2)
  Printer.Line (45, 13)-(45, Y2)
  Printer.Line (65, 13)-(65, Y2)
  Printer.Line (75, 13)-(75, Y2)
  Printer.Line (95, 13)-(95, Y2)
  Printer.Line (105, 13)-(105, Y2)
  Printer.Line (125, 13)-(125, Y2)
  Printer.Line (135, 13)-(135, Y2)
  Printer.Line (155, 13)-(155, Y2)
  Printer.Line (165, 13)-(165, Y2)
  Printer.Line (Largura, 13)-(Largura, Y2)
  
  Printer.EndDoc
End With
naoImprime:

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
If dbContagem.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField Then
  dbContagem.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField & " desc"
Else
  dbContagem.Recordset.Sort = DataGrid1.Columns(ColIndex).DataField
End If
End Sub

Private Sub dbContagem_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
If dbContagem.Recordset.RecordCount = 0 Then Exit Sub
If dbContagem.Recordset.EOF = True Then Exit Sub
With dbDepositados
  .ConnectionString = CaminhoADO
  .RecordSource = "select count(codigocheque) as Cheques, sum(valor) as Total from cheques where compensado=-1 and banco='" & dbContagem.Recordset!debanco & "' and agencia='" & dbContagem.Recordset!deAgencia & "' and conta='" & dbContagem.Recordset!deconta & "' and cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
End With
With dbDevolvidos
  .ConnectionString = CaminhoADO
  .RecordSource = "select count(codigocheque) as Cheques, sum(valor) as Total from cheques where devolvido=-1 and cobrando=0 and banco='" & dbContagem.Recordset!debanco & "' and agencia='" & dbContagem.Recordset!deAgencia & "' and conta='" & dbContagem.Recordset!deconta & "' and cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
End With
With dbCobrando
  .ConnectionString = CaminhoADO
  .RecordSource = "select count(codigocheque) as Cheques, sum(valor) as Total from cheques where devolvido=-1 and cobrando=-1 and banco='" & dbContagem.Recordset!debanco & "' and agencia='" & dbContagem.Recordset!deAgencia & "' and conta='" & dbContagem.Recordset!deconta & "' and cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
End With
With dbPre
  .ConnectionString = CaminhoADO
  .RecordSource = "select count(codigocheque) as Cheques, sum(valor) as Total from cheques where devolvido=0 and cobrando=0 and compensado=0 and banco='" & dbContagem.Recordset!debanco & "' and agencia='" & dbContagem.Recordset!deAgencia & "' and conta='" & dbContagem.Recordset!deconta & "' and cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "#"
  .Refresh
End With

End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = DateAdd("m", 1, Date)

With dbContagem
  .ConnectionString = CaminhoADO
  .RecordSource = "SELECT Count(Cheques.CodigoFechamento) AS Cheques, Sum(Cheques.Valor) AS total, First(Cheques.Banco) AS DeBanco, First(Cheques.Agencia) AS DeAgencia, First(Cheques.Conta) AS DeConta From Cheques where cheques.datacheque between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# GROUP BY cheques.banco, cheques.agencia, cheques.conta"
  .Refresh
  .Recordset.Sort = "debanco, deagencia, deconta"
End With

End Sub
