VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmDespesasConfirmaExibePg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagamentos da Despesa"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6990
   Icon            =   "frmDespesasConfirmaExibePg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data dbPagamentos 
      Caption         =   "dbPagamentos"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Meus documentos\Fabio\Projeto For Windows\Posto\Posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from QConciliaNovaContas where tipo='Despesa' and codigo=0"
      Top             =   960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdSair 
      Cancel          =   -1  'True
      Caption         =   "Sair"
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmDespesasConfirmaExibePg.frx":0442
      Height          =   2055
      Left            =   0
      OleObjectBlob   =   "frmDespesasConfirmaExibePg.frx":045D
      TabIndex        =   1
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Pago:"
      Height          =   195
      Left            =   4320
      TabIndex        =   3
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label lblTotalPago 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5280
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
End
Attribute VB_Name = "frmDespesasConfirmaExibePg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Exibir(ByVal CodigoDespesa As Double)
Dim TempValor As Currency
With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as pago from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=" & CodigoDespesa
  .Refresh
  If IsNull(.Recordset!Pago) = False Then
    TempValor = .Recordset!Pago
  Else
    TempValor = 0
  End If
  .RecordSource = "select *from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=" & CodigoDespesa
  .Refresh
End With
lblTotalPago.Caption = Format(TempValor, "Currency")

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
With dbPagamentos
  .Connect = Conectar
  .DatabaseName = Caminho
  .RecordSource = "select sum(valor) as pago from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=0"
  .Refresh
  If IsNull(.Recordset!Pago) = False Then
    TempValor = .Recordset!Pago
  Else
    TempValor = 0
  End If
  .RecordSource = "select *from QConciliaNovaContas where concilianova.tipo='Despesa' and codigo=0"
  .Refresh
End With

End Sub
