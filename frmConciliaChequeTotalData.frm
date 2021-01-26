VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmConciliaChequeTotalData 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total de Cheques por Data"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   Icon            =   "frmConciliaChequeTotalData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.Data QDevolvidos 
      Caption         =   "QDevolvidos"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   248
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmConciliaChequeTotalData.frx":0442
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QSomaDevolvidos 
      Caption         =   "QSomaDevolvidos"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   248
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where compensado=0 and devolvido=-1 and cobrando=0 and protesto=0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QSomaCobra 
      Caption         =   "QSomaCobra"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3368
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmConciliaChequeTotalData.frx":04CF
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QSomaCobraTotal 
      Caption         =   "QSomaCobraTotal"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3368
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where compensado=0 and cobrando=-1 and devolvido=-1 and protesto=0"
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QChequesCustodiaTotal 
      Caption         =   "QChequesCustodiaTotal"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3368
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from compensapendente where conciliado=0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QChequesCustodia 
      Caption         =   "QChequesCustodia"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3368
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select *from compensapendente where conciliado=0 order by data"
      Top             =   840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QSomaChequesTotal 
      Caption         =   "QSomaChequesTotal"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   248
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "select sum(valor) as total from cheques where codigosoma='1' and compensado=0 and cobrando=0 and devolvido=0 and protesto=0"
      Top             =   1080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data QSomaCheques 
      Caption         =   "QSomaCheques"
      Connect         =   "Access"
      DatabaseName    =   "posto.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   248
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   $"frmConciliaChequeTotalData.frx":0565
      Top             =   720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmConciliaChequeTotalData.frx":0604
      Height          =   2175
      Left            =   128
      OleObjectBlob   =   "frmConciliaChequeTotalData.frx":061F
      TabIndex        =   1
      Top             =   0
      Width           =   3135
   End
   Begin VB.CommandButton cmdOk 
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Height          =   375
      Left            =   2648
      TabIndex        =   0
      Top             =   5520
      Width           =   1215
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frmConciliaChequeTotalData.frx":1026
      Height          =   2175
      Left            =   3248
      OleObjectBlob   =   "frmConciliaChequeTotalData.frx":1045
      TabIndex        =   4
      Top             =   0
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmConciliaChequeTotalData.frx":1A48
      Height          =   2175
      Left            =   128
      OleObjectBlob   =   "frmConciliaChequeTotalData.frx":1A62
      TabIndex        =   7
      Top             =   2760
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid4 
      Bindings        =   "frmConciliaChequeTotalData.frx":2469
      Height          =   2175
      Left            =   3248
      OleObjectBlob   =   "frmConciliaChequeTotalData.frx":2482
      TabIndex        =   8
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label lblDevolvidos 
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
      Left            =   1448
      TabIndex        =   12
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   488
      TabIndex        =   11
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label lblCobra 
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
      Left            =   4568
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   3608
      TabIndex        =   9
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   3608
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblCustodia 
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
      Left            =   4568
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total:"
      Height          =   255
      Left            =   488
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblPre 
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
      Left            =   1448
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "frmConciliaChequeTotalData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
With qSomaCheques
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QSomaChequesTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblPre.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblPre.Caption = Format(0, "Currency")
  End If
End With

With QDevolvidos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QSomaDevolvidos
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblDevolvidos.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblDevolvidos.Caption = Format(0, "Currency")
  End If
End With

With QSomaCobra
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QSomaCobraTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblCobra.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblCobra.Caption = Format(0, "Currency")
  End If
End With

With QChequesCustodia
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
End With
With QChequesCustodiaTotal
  .Connect = Conectar
  .DatabaseName = Caminho
  .Refresh
  If IsNull(.Recordset!Total) = False Then
    lblCustodia.Caption = Format(.Recordset!Total, "Currency")
  Else
    lblCustodia.Caption = Format(0, "Currency")
  End If
End With

End Sub

