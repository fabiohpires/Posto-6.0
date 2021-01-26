VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaChequesDepositados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cheques depositados"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dbCompensaPendente 
      Caption         =   "dbCompensaPendente"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CompensaPendente"
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Imprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Data dbChequesTotal 
      Caption         =   "dbChequesTotal"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cheques"
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data dbCheques 
      Caption         =   "dbCheques"
      Connect         =   "Access 2000;"
      DatabaseName    =   "D:\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cheques"
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaChequesDepositados.frx":0000
      Height          =   5895
      Left            =   120
      OleObjectBlob   =   "frmConciliaChequesDepositados.frx":0018
      TabIndex        =   0
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   6120
      Width           =   1935
   End
End
Attribute VB_Name = "frmConciliaChequesDepositados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tipo As String, Dia  As Date, Valor As Currency, NrDocumento As String

Public Sub Exibe()
If Tipo = "Custódia de cheques!" Then
  With dbCompensaPendente
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from compensapendente where data=#" & DataInglesa(Dia) & "# and valor=" & NumeroIngles(Valor)
    .Refresh
    If .Recordset.RecordCount <> 0 Then
      'Previsao = .Recordset!CodigoPendencia
      Previsao = .Recordset!NrDoc
    End If
  End With
  'On Error Resume Next
  With dbCheques
    .Connect = Conectar
    .DatabaseName = Caminho
    '.RecordSource = "select *from cheques where codigoprevisaorecebe=" & Previsao
    .RecordSource = "select *from cheques where datacomp=#" & DataInglesa(dbCompensaPendente.Recordset!Data) & "# and codigosoma='" & Previsao & "'"
    .Refresh
  End With
  
  With dbChequesTotal
    .Connect = Conectar
    .DatabaseName = Caminho
    '.RecordSource = "select sum(valor) as total from cheques where codigoprevisaorecebe=" & Previsao
    .RecordSource = "select sum(valor) as total from cheques where datacomp=#" & DataInglesa(dbCompensaPendente.Recordset!Data) & "# and codigosoma='" & Previsao & "'"
    .Refresh
    If IsNull(.Recordset!Total) = False Then
      lblTotal.Caption = Format(.Recordset!Total, "currency")
    Else
      lblTotal.Caption = Format(0, "currency")
    End If
  End With
Else
  With dbCheques
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select *from cheques where codigosoma='" & NrDocumento & "'"
    .Refresh
  End With
  
  With dbChequesTotal
    .Connect = Conectar
    .DatabaseName = Caminho
    .RecordSource = "select sum(valor) as total from cheques where codigosoma='" & NrDocumento & "'"
    .Refresh
    If IsNull(.Recordset!Total) = False Then
      lblTotal.Caption = Format(.Recordset!Total, "currency")
    Else
      lblTotal.Caption = Format(0, "currency")
    End If
  End With
End If

End Sub

Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Previsao
With dbCompensaPendente
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
End Sub

Private Sub Imprimir_Click()
  On Error GoTo NaoImprime
  If ShowPrinter(Me) = 0 Then Exit Sub
  On Error GoTo 0
  ImprimeGrid DBGrid1, Printer, dbCheques, 6, True, , , , , "Cheques Depositados", NomePosto, "Data do depósito: " & Format(frmConciliacaoNova.dbConcilia.Recordset!Data, "long Date")
  
  Printer.EndDoc
NaoImprime:
End Sub
