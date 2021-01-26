VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmRelatTotalVendido 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total Vendido"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   Begin VB.Data dbProdutos 
      Caption         =   "dbProdutos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Produtos"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.Data dbVendas 
      Caption         =   "dbVendas"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "VendasTotalTemp"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data dbTemp 
      Caption         =   "dbTemp"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmTotalVendido.frx":0000
      Height          =   3015
      Left            =   120
      OleObjectBlob   =   "frmTotalVendido.frx":0017
      TabIndex        =   5
      Top             =   840
      Width           =   4215
   End
   Begin VB.CommandButton cmdExibir 
      Caption         =   "Exibir"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin MSComCtl2.DTPicker txtDataIni 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   37665
   End
   Begin MSComCtl2.DTPicker txtDataFim 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Format          =   58916865
      CurrentDate     =   37665
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "a"
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   90
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Período:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmRelatTotalVendido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdExibir_Click()
Dim Total As Currency
With dbVendas
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
With dbTemp
'  .RecordSource = "select sum(valorvendido) as total, codigoproduto from bicomovimento where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by codigoproduto"
'  .Refresh
'  If .Recordset.RecordCount <> 0 Then
'    .Recordset.MoveLast
'    .Recordset.MoveFirst
'    Do While .Recordset.EOF = False
'      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
'      If dbProdutos.Recordset.NoMatch = False Then
'        If dbVendas.Recordset.RecordCount = 0 Then
'          dbVendas.Recordset.AddNew
'          dbVendas.Recordset!totalvendido = 0
'        Else
'          dbVendas.Recordset.FindFirst "descricao='" & dbProdutos.Recordset!categoria & "'"
'          If dbVendas.Recordset.NoMatch = False Then
'            dbVendas.Recordset.Edit
'          Else
'            dbVendas.Recordset.AddNew
'            dbVendas.Recordset!totalvendido = 0
'          End If
'        End If
'      Else
'        dbVendas.Recordset.AddNew
'        dbVendas.Recordset!totalvendido = 0
'      End If
'      Total = Total + .Recordset!Total
'      dbVendas.Recordset!Descricao = dbProdutos.Recordset!categoria
'      dbVendas.Recordset!totalvendido = dbVendas.Recordset!totalvendido + .Recordset!Total
'      dbVendas.Recordset.Update
'      .Recordset.MoveNext
'    Loop
'  End If
  
  
  .RecordSource = "select sum(valortotal) as total, bicoencerrantes.codigoproduto from qbicoencerrantes where datacaixa between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by bicoencerrantes.codigoproduto"
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbProdutos.Recordset.FindFirst "codigoproduto=" & .Recordset!CodigoProduto
      If dbProdutos.Recordset.NoMatch = False Then
        If dbVendas.Recordset.RecordCount = 0 Then
          dbVendas.Recordset.AddNew
          dbVendas.Recordset!totalvendido = 0
        Else
          dbVendas.Recordset.FindFirst "descricao='" & dbProdutos.Recordset!categoria & "'"
          If dbVendas.Recordset.NoMatch = False Then
            dbVendas.Recordset.Edit
          Else
            dbVendas.Recordset.AddNew
            dbVendas.Recordset!totalvendido = 0
          End If
        End If
      Else
        dbVendas.Recordset.AddNew
        dbVendas.Recordset!totalvendido = 0
      End If
      Total = Total + .Recordset!Total
      dbVendas.Recordset!Descricao = dbProdutos.Recordset!categoria
      dbVendas.Recordset!totalvendido = dbVendas.Recordset!totalvendido + .Recordset!Total
      dbVendas.Recordset.Update
      .Recordset.MoveNext
    Loop
  End If
  
  
  
  .RecordSource = "select sum(valor) as total, categoria from qvendadiaprodutos where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by categoria"
  .Refresh
  
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbVendas.Recordset.FindFirst "descricao='" & .Recordset!categoria & "'"
      If dbVendas.Recordset.NoMatch = False Then
        dbVendas.Recordset.Edit
      Else
        dbVendas.Recordset.AddNew
        dbVendas.Recordset!totalvendido = 0
      End If
      Total = Total + .Recordset!Total
      dbVendas.Recordset!Descricao = .Recordset!categoria
      dbVendas.Recordset!totalvendido = dbVendas.Recordset!totalvendido + .Recordset!Total
      dbVendas.Recordset.Update
      
      .Recordset.MoveNext
    Loop
  End If
  
  
  .RecordSource = "select sum(totalvendido) as total, categoria from qvendadiaprodutos2 where data between #" & DataInglesa(Trim(Str(txtDataIni.Value))) & "# and #" & DataInglesa(Trim(Str(txtDataFim.Value))) & "# group by categoria"
  .Refresh
  
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      dbVendas.Recordset.FindFirst "descricao='" & .Recordset!categoria & "'"
      If dbVendas.Recordset.NoMatch = False Then
        dbVendas.Recordset.Edit
      Else
        dbVendas.Recordset.AddNew
        dbVendas.Recordset!totalvendido = 0
      End If
      Total = Total + .Recordset!Total
      dbVendas.Recordset!Descricao = .Recordset!categoria
      dbVendas.Recordset!totalvendido = dbVendas.Recordset!totalvendido + .Recordset!Total
      dbVendas.Recordset.Update
      
      .Recordset.MoveNext
    Loop
  End If
  
  
  
  lblTotal.Caption = Format(Total, "Currency")
End With
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtDataIni.Value = DateAdd("m", -1, Date)
txtDataFim.Value = Date
With dbVendas
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If .Recordset.RecordCount <> 0 Then
    .Recordset.MoveLast
    .Recordset.MoveFirst
    Do While .Recordset.EOF = False
      .Recordset.Delete
      .Refresh
    Loop
  End If
End With
With dbTemp
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With dbProdutos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

